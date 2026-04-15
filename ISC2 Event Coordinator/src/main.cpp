// ISC2 Event Coordinator - Console Application
// Processes Constant Contact attendee reports and generates attendance sheets and name tags.

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <winhttp.h>
#include <wincrypt.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <comdef.h>
#include <shellapi.h>

#include <pugixml.hpp>
#include <nlohmann/json.hpp>

#include <algorithm>
#include <cctype>
#include <cstdio>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <map>
#include <set>
#include <sstream>
#include <string>
#include <vector>

namespace fs = std::filesystem;
using json = nlohmann::json;

// ============================================================================
// Utility functions
// ============================================================================

static std::wstring utf8_to_wide(const std::string& s) {
    if (s.empty()) return {};
    int len = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
    std::wstring ws(len, 0);
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), ws.data(), len);
    return ws;
}

static std::string wide_to_utf8(const std::wstring& ws) {
    if (ws.empty()) return {};
    int len = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), nullptr, 0, nullptr, nullptr);
    std::string s(len, 0);
    WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), s.data(), len, nullptr, nullptr);
    return s;
}

static std::string capitalize(const std::string& s) {
    std::string result = s;
    bool capitalize_next = true;
    for (auto& c : result) {
        if (capitalize_next && std::isalpha(static_cast<unsigned char>(c))) {
            c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
            capitalize_next = false;
        } else if (c == ' ' || c == '-' || c == '\'') {
            capitalize_next = true;
        } else {
            c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
        }
    }
    return result;
}

static std::string trim(const std::string& s) {
    auto start = s.find_first_not_of(" \t\r\n");
    if (start == std::string::npos) return "";
    auto end = s.find_last_not_of(" \t\r\n");
    return s.substr(start, end - start + 1);
}

// Parse "Month DD, YYYY" from the beginning of an event name and return "YYYY-MM-DD".
// Returns empty string if the name doesn't start with a recognisable date.
static std::string parse_event_date(const std::string& name) {
    static const char* months[] = {
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    };
    for (int m = 0; m < 12; ++m) {
        std::string prefix = months[m];
        if (name.size() > prefix.size() && name.substr(0, prefix.size()) == prefix
            && name[prefix.size()] == ' ') {
            // Expect "DD, YYYY" after the month + space
            size_t comma = name.find(',', prefix.size() + 1);
            if (comma == std::string::npos) break;
            std::string day_str = trim(name.substr(prefix.size() + 1, comma - prefix.size() - 1));
            std::string year_str = trim(name.substr(comma + 1, 5)); // " YYYY"
            int day = 0, year = 0;
            try { day = std::stoi(day_str); year = std::stoi(year_str); } catch (...) { break; }
            if (year < 2000 || year > 2099 || day < 1 || day > 31) break;
            char buf[16];
            snprintf(buf, sizeof(buf), "%04d-%02d-%02d", year, m + 1, day);
            return std::string(buf);
        }
    }
    return "";
}

// ============================================================================
// CSV Parser - handles quoted fields with commas and embedded quotes
// ============================================================================

struct CsvRow {
    std::vector<std::string> fields;
    std::string get(size_t idx) const {
        if (idx < fields.size()) return fields[idx];
        return "";
    }
};

// Validate that a CSV file has the expected Constant Contact "Attendee report" header.
static bool is_valid_attendee_csv(const std::string& filepath) {
    std::ifstream file(filepath);
    if (!file.is_open()) return false;
    std::string header_line;
    if (!std::getline(file, header_line)) return false;
    // The header must contain these key columns from the Constant Contact attendee report
    return header_line.find("Ticket Type") != std::string::npos
        && header_line.find("Attendee's First Name") != std::string::npos
        && header_line.find("Attendee's Last Name") != std::string::npos
        && header_line.find("Attendee's Email Address") != std::string::npos;
}

static std::vector<CsvRow> parse_csv(const std::string& filepath) {
    std::vector<CsvRow> rows;
    std::ifstream file(filepath);
    if (!file.is_open()) {
        std::cerr << "Error: Cannot open CSV file: " << filepath << "\n";
        return rows;
    }

    std::string line;
    while (std::getline(file, line)) {
        // Remove trailing \r if present
        if (!line.empty() && line.back() == '\r') line.pop_back();

        CsvRow row;
        std::string field;
        bool in_quotes = false;

        for (size_t i = 0; i < line.size(); ++i) {
            char c = line[i];
            if (in_quotes) {
                if (c == '"') {
                    if (i + 1 < line.size() && line[i + 1] == '"') {
                        field += '"';
                        ++i;
                    } else {
                        in_quotes = false;
                    }
                } else {
                    field += c;
                }
            } else {
                if (c == '"') {
                    in_quotes = true;
                } else if (c == ',') {
                    row.fields.push_back(trim(field));
                    field.clear();
                } else {
                    field += c;
                }
            }
        }
        row.fields.push_back(trim(field));

        if (!row.fields.empty() && !(row.fields.size() == 1 && row.fields[0].empty())) {
            rows.push_back(std::move(row));
        }
    }
    return rows;
}

// ============================================================================
// Constant Contact API integration
// ============================================================================

// Token file path and executable name (set from main() before download_attendee_report is called)
static fs::path g_token_file_path;
static std::string g_exe_name;

// Error log path (set from run() after base directory is determined)
static fs::path g_errors_log_path;

static void write_error_log(int error_num, const std::string& description, const std::string& detail) {
    if (g_errors_log_path.empty()) return;
    SYSTEMTIME st;
    GetLocalTime(&st);
    char time_buf[32];
    snprintf(time_buf, sizeof(time_buf), "%04d-%02d-%02dT%02d:%02d:%02d",
             st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    std::ofstream log(g_errors_log_path, std::ios::app);
    if (!log.is_open()) return;
    char nn[4];
    snprintf(nn, sizeof(nn), "%02d", error_num);
    if (description.empty())
        log << "Error" << nn << ":\n";
    else
        log << "Error" << nn << ": " << description << "\n";
    log << "Time: " << time_buf << "\n";
    log << "Details:\n" << detail << "\n";
    log << "\n";
}

struct ConstantContactTokens {
    std::string api_key;
    std::string client_secret;
    std::string access_token;
    std::string refresh_token;
};

static bool save_tokens(const ConstantContactTokens& tok) {
    if (g_token_file_path.empty()) return false;
    json j;
    j["api_key"] = tok.api_key;
    j["client_secret"] = tok.client_secret;
    j["access_token"] = tok.access_token;
    j["refresh_token"] = tok.refresh_token;
    std::string plain = j.dump(2) + "\n";
    // Encrypt with DPAPI (current-user scope)
    DATA_BLOB in_blob, out_blob;
    in_blob.pbData = reinterpret_cast<BYTE*>(plain.data());
    in_blob.cbData = static_cast<DWORD>(plain.size());
    if (!CryptProtectData(&in_blob, L"ISC2TokenFile", nullptr, nullptr, nullptr, 0, &out_blob)) {
        // DPAPI unavailable — write plaintext as fallback
        std::ofstream f(g_token_file_path);
        if (!f.is_open()) return false;
        f << plain;
        return true;
    }
    std::ofstream f(g_token_file_path, std::ios::binary);
    bool ok = f.is_open();
    if (ok) f.write(reinterpret_cast<const char*>(out_blob.pbData), out_blob.cbData);
    LocalFree(out_blob.pbData);
    return ok;
}

static ConstantContactTokens load_tokens() {
    ConstantContactTokens tok;
    if (g_token_file_path.empty() || !fs::exists(g_token_file_path)) return tok;
    // Read file as binary (works for both DPAPI blob and legacy plaintext)
    std::ifstream f(g_token_file_path, std::ios::binary);
    if (!f.is_open()) return tok;
    std::vector<BYTE> file_data((std::istreambuf_iterator<char>(f)),
                                 std::istreambuf_iterator<char>());
    f.close();
    std::string plain;
    // Try DPAPI decryption first
    DATA_BLOB in_blob, out_blob;
    in_blob.pbData = file_data.data();
    in_blob.cbData = static_cast<DWORD>(file_data.size());
    if (!file_data.empty() &&
        CryptUnprotectData(&in_blob, nullptr, nullptr, nullptr, nullptr, 0, &out_blob)) {
        plain.assign(reinterpret_cast<const char*>(out_blob.pbData), out_blob.cbData);
        LocalFree(out_blob.pbData);
    } else {
        // Fallback: treat as plaintext (migration from pre-DPAPI format)
        plain.assign(file_data.begin(), file_data.end());
    }
    try {
        json j = json::parse(plain);
        tok.api_key       = j.value("api_key", "");
        tok.client_secret = j.value("client_secret", "");
        tok.access_token  = j.value("access_token", "");
        tok.refresh_token = j.value("refresh_token", "");
    } catch (...) {}
    return tok;
}

static bool refresh_access_token(ConstantContactTokens& tok) {
    if (tok.refresh_token.empty() || tok.api_key.empty() || tok.client_secret.empty()) {
        std::cerr << "Cannot refresh token: missing";
        if (tok.refresh_token.empty()) std::cerr << " refresh_token";
        if (tok.client_secret.empty()) std::cerr << " client_secret";
        if (tok.api_key.empty())       std::cerr << " api_key";
        std::cerr << " in tokens.json\n";
        return false;
    }
    std::cout << "Access token expired. Refreshing...\n";

    // POST to authz.constantcontact.com/oauth2/default/v1/token
    // with Basic auth (api_key:client_secret) and grant_type=refresh_token
    std::string credentials = tok.api_key + ":" + tok.client_secret;
    // Base64 encode
    DWORD b64_len = 0;
    CryptBinaryToStringA((const BYTE*)credentials.c_str(), (DWORD)credentials.size(),
                         CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, nullptr, &b64_len);
    std::string b64(b64_len, '\0');
    CryptBinaryToStringA((const BYTE*)credentials.c_str(), (DWORD)credentials.size(),
                         CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, b64.data(), &b64_len);
    b64.resize(b64_len);
    // Trim trailing nulls
    while (!b64.empty() && b64.back() == '\0') b64.pop_back();

    std::string body = "grant_type=refresh_token&refresh_token=" + tok.refresh_token;
    std::string auth_header = "Basic " + b64;

    // Use WinHTTP to POST
    HINTERNET hSession = WinHttpOpen(L"ISC2EventCoordinator/1.0",
                                     WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                     WINHTTP_NO_PROXY_NAME,
                                     WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return false;

    HINTERNET hConnect = WinHttpConnect(hSession, L"authz.constantcontact.com",
                                        INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST",
                                            L"/oauth2/default/v1/token",
                                            nullptr, WINHTTP_NO_REFERER,
                                            WINHTTP_DEFAULT_ACCEPT_TYPES,
                                            WINHTTP_FLAG_SECURE);
    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    std::wstring auth_w = L"Authorization: " + utf8_to_wide(auth_header);
    WinHttpAddRequestHeaders(hRequest, auth_w.c_str(), (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);
    WinHttpAddRequestHeaders(hRequest, L"Content-Type: application/x-www-form-urlencoded",
                             (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);

    std::string response;
    if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                           (LPVOID)body.c_str(), (DWORD)body.size(),
                           (DWORD)body.size(), 0) &&
        WinHttpReceiveResponse(hRequest, nullptr)) {
        DWORD size = 0;
        do {
            WinHttpQueryDataAvailable(hRequest, &size);
            if (size > 0) {
                std::vector<char> buf(size + 1, 0);
                DWORD downloaded = 0;
                WinHttpReadData(hRequest, buf.data(), size, &downloaded);
                response.append(buf.data(), downloaded);
            }
        } while (size > 0);
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    if (response.empty()) {
        std::cerr << "Token refresh failed: empty response.\n";
        return false;
    }

    try {
        json resp = json::parse(response);
        if (resp.contains("access_token")) {
            tok.access_token  = resp["access_token"].get<std::string>();
            if (resp.contains("refresh_token")) {
                tok.refresh_token = resp["refresh_token"].get<std::string>();
            }
            save_tokens(tok);
            std::cout << "Token refreshed successfully.\n";
            return true;
        } else {
            std::cerr << "Token refresh error: " << response.substr(0, 300) << "\n";
            return false;
        }
    } catch (...) {
        std::cerr << "Token refresh: failed to parse response.\n";
        return false;
    }
}

static std::string http_get(const std::wstring& host, const std::wstring& path,
                            const std::string& bearer_token,
                            const std::string& api_key = "") {
    std::string result;
    HINTERNET hSession = WinHttpOpen(L"ISC2EventCoordinator/1.0",
                                     WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                     WINHTTP_NO_PROXY_NAME,
                                     WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return result;

    HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(),
                                        INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return result; }

    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(),
                                            nullptr, WINHTTP_NO_REFERER,
                                            WINHTTP_DEFAULT_ACCEPT_TYPES,
                                            WINHTTP_FLAG_SECURE);
    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return result;
    }

    std::wstring auth = L"Authorization: Bearer " + utf8_to_wide(bearer_token);
    WinHttpAddRequestHeaders(hRequest, auth.c_str(), (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);

    if (!api_key.empty()) {
        std::wstring key_hdr = L"x-api-key: " + utf8_to_wide(api_key);
        WinHttpAddRequestHeaders(hRequest, key_hdr.c_str(), (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);
    }

    if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                           WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
        WinHttpReceiveResponse(hRequest, nullptr)) {
        // Log HTTP status code
        DWORD status_code = 0;
        DWORD sc_size = sizeof(status_code);
        WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                            WINHTTP_HEADER_NAME_BY_INDEX, &status_code, &sc_size,
                            WINHTTP_NO_HEADER_INDEX);
        if (status_code != 200) {
            std::cerr << "  HTTP " << status_code << " from GET "
                      << wide_to_utf8(path) << "\n";
        }

        DWORD size = 0;
        do {
            WinHttpQueryDataAvailable(hRequest, &size);
            if (size > 0) {
                std::vector<char> buf(size + 1, 0);
                DWORD downloaded = 0;
                WinHttpReadData(hRequest, buf.data(), size, &downloaded);
                result.append(buf.data(), downloaded);
            }
        } while (size > 0);
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);
    return result;
}

static std::string http_post(const std::wstring& host, const std::wstring& path,
                             const std::string& bearer_token,
                             const std::string& content_type,
                             const std::string& body) {
    std::string result;
    HINTERNET hSession = WinHttpOpen(L"ISC2EventCoordinator/1.0",
                                     WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                     WINHTTP_NO_PROXY_NAME,
                                     WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return result;

    HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(),
                                        INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return result; }

    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", path.c_str(),
                                            nullptr, WINHTTP_NO_REFERER,
                                            WINHTTP_DEFAULT_ACCEPT_TYPES,
                                            WINHTTP_FLAG_SECURE);
    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return result;
    }

    std::wstring auth = L"Authorization: Bearer " + utf8_to_wide(bearer_token);
    WinHttpAddRequestHeaders(hRequest, auth.c_str(), (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);

    std::wstring ct = L"Content-Type: " + utf8_to_wide(content_type);
    WinHttpAddRequestHeaders(hRequest, ct.c_str(), (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);

    if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                           (LPVOID)body.c_str(), (DWORD)body.size(),
                           (DWORD)body.size(), 0) &&
        WinHttpReceiveResponse(hRequest, nullptr)) {
        DWORD size = 0;
        do {
            WinHttpQueryDataAvailable(hRequest, &size);
            if (size > 0) {
                std::vector<char> buf(size + 1, 0);
                DWORD downloaded = 0;
                WinHttpReadData(hRequest, buf.data(), size, &downloaded);
                result.append(buf.data(), downloaded);
            }
        } while (size > 0);
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);
    return result;
}

static std::string csv_escape(const std::string& field) {
    if (field.find(',') != std::string::npos || field.find('"') != std::string::npos
        || field.find('\n') != std::string::npos) {
        std::string escaped = "\"";
        for (char c : field) {
            if (c == '"') escaped += "\"\"";
            else escaped += c;
        }
        escaped += '"';
        return escaped;
    }
    return field;
}

static bool download_attendee_report(fs::path& output_csv_path,
                                     std::string& out_event_title,
                                     std::string& out_event_date) {
    std::cout << "\n=== Step 1: Download Attendee Report from Constant Contact ===\n\n";


    // Determine the output directory: if output_csv_path is a directory, use it; if it's a file, use its parent
    fs::path output_dir;
    if (fs::is_directory(output_csv_path)) {
        output_dir = output_csv_path;
    } else {
        output_dir = output_csv_path.parent_path();
    }

    std::vector<fs::path> csv_files;
    if (!output_dir.empty() && fs::exists(output_dir)) {
        for (auto& entry : fs::directory_iterator(output_dir)) {
            if (entry.is_regular_file() && entry.path().extension() == ".csv") {
                // Only include files that pass Constant Contact Attendee report validation
                if (!is_valid_attendee_csv(entry.path().string())) continue;
                csv_files.push_back(entry.path());
            }
        }
        std::sort(csv_files.begin(), csv_files.end());
    }

    if (!csv_files.empty()) {
        std::cout << "Existing Constant Contact 'Attendee report' CSV files found in the output folder:\n";
        for (size_t i = 0; i < csv_files.size(); ++i) {
            std::cout << "  " << (i + 1) << ". " << csv_files[i].filename().string() << "\n";
        }
        std::cout << "  " << (csv_files.size() + 1) << ". Download new data from Constant Contact API\n";
        int csv_choice = 0;
        while (true) {
            std::cout << "\nChoose an option (enter number): ";
            std::string response;
            std::getline(std::cin, response);
            { std::string t = trim(response); csv_choice = (!t.empty() && std::all_of(t.begin(), t.end(), ::isdigit)) ? std::stoi(t) : 0; }
            if (csv_choice >= 1 && csv_choice <= (int)(csv_files.size() + 1)) break;
            std::cout << "Invalid Response\n";
        }
        if (csv_choice != (int)(csv_files.size() + 1)) {
            output_csv_path = csv_files[csv_choice - 1];
            std::cout << "Using: " << output_csv_path.filename().string() << "\n";
            // Extract prefix: everything before "Registration Spreadsheet.csv" in the filename.
            // e.g. "2026-03-24.Registration Spreadsheet.csv" -> "2026-03-24."
            //      "test.Registration Spreadsheet.csv"     -> "test."
            //      "Registration Spreadsheet.csv"          -> ""
            std::string fname = output_csv_path.filename().string();
            const std::string marker = "Registration Spreadsheet.csv";
            size_t marker_pos = fname.find(marker);
            if (marker_pos != std::string::npos && marker_pos > 0) {
                // Strip the trailing dot so out_event_date = "2026-03-24" or "test"
                out_event_date = fname.substr(0, marker_pos - 1); // remove trailing '.'
            }
            // out_event_date remains empty if the file is plain "Registration Spreadsheet.csv"
            return true;
        }
    }

    std::cout << "To download from Constant Contact API, you need two credentials:\n";
    std::cout << "  1) API Key (Client ID)\n";
    std::cout << "  2) OAuth2 Access Token\n\n";

    // Load saved tokens
    ConstantContactTokens tokens = load_tokens();
    std::string api_key, token;

    if (!tokens.api_key.empty() && !tokens.access_token.empty()) {
        std::cout << "Saved credentials found (API Key: " << tokens.api_key.substr(0, 8) << "...).\n";
        std::string use_saved;
        while (true) {
            std::cout << "Use saved credentials? (y/n): ";
            std::getline(std::cin, use_saved);
            use_saved = trim(use_saved);
            if (use_saved.size() == 1 && (use_saved[0] == 'y' || use_saved[0] == 'Y' ||
                                          use_saved[0] == 'n' || use_saved[0] == 'N')) break;
            std::cout << "Invalid Response\n";
        }
        if (use_saved.size() == 1 && (use_saved[0] == 'y' || use_saved[0] == 'Y')) {
            api_key = tokens.api_key;
            token   = tokens.access_token;
        }
    }

    if (api_key.empty()) {
        if (!fs::exists(g_token_file_path))
            std::cout << "File " << g_token_file_path.string() << " was not found.\n";
        std::cout << "Enter your API Key / Client ID (or press Enter to skip API download): ";
        std::getline(std::cin, api_key);
        api_key = trim(api_key);
    }

    if (api_key.empty()) {
        write_error_log(2, "No API key provided and no CSV file available.",
            "User did not provide an API key and no valid CSV file exists in the output folder.");
        std::cerr << "\nNo CSV file available and no API key provided.\n";
        std::cerr << "Cannot continue without a registration spreadsheet.\n";
        return false;
    }

    if (token.empty()) {
        std::cout << "Enter your OAuth2 Access Token: ";
        std::getline(std::cin, token);
        token = trim(token);
    }

    if (token.empty()) {
        write_error_log(3, "No access token provided and no CSV file available.",
            "User did not provide an OAuth2 access token and no valid CSV file exists in the output folder.");
        std::cerr << "\nNo access token provided and no CSV file available.\n";
        std::cerr << "Cannot continue without a registration spreadsheet.\n";
        return false;
    }

    // Prompt for client secret if not saved (needed for token refresh)
    if (tokens.client_secret.empty()) {
        std::cout << "Enter your Client Secret (for automatic token refresh, or Enter to skip): ";
        std::string secret;
        std::getline(std::cin, secret);
        tokens.client_secret = trim(secret);
    }

    // Save tokens for future runs
    tokens.api_key = api_key;
    tokens.access_token = token;
    save_tokens(tokens);

    // --- Fetch list of events from the Events API ---
    // Include all statuses so completed/ended events appear.
    std::cout << "Fetching events from Constant Contact...\n";
    std::string events_json = http_get(
        L"api.cc.email",
        L"/v3/events?limit=50",
        token, api_key);

    // Auto-refresh on 401 (response may say "invalid_token" or "unauthorized")
    bool auth_failed = false;
    if (events_json.find("unauthorized") != std::string::npos
        || events_json.find("Unauthorized") != std::string::npos
        || events_json.find("invalid_token") != std::string::npos) {
        if (refresh_access_token(tokens)) {
            token = tokens.access_token;
            events_json = http_get(
                L"api.cc.email",
                L"/v3/events?limit=50",
                token, api_key);
        } else {
            auth_failed = true;
        }
    }

    // --- API failure handler: write to messages\errors.log and display structured instructions ---
    auto api_failure = [&](const std::string& error_detail) -> bool {
        write_error_log(1, "Accessing the Constant Contact service failed.", error_detail);
        std::string out_folder = output_dir.string();
        std::string log_ref = g_errors_log_path.empty()
            ? (out_folder + "\\..\\messages\\errors.log")
            : g_errors_log_path.string();
        std::cerr << "\nError01: Accessing the Constant Contact service failed.\n"
                  << "How to Proceed:\n"
                  << "Step 1: Login to Constant Contact (https://login.constantcontact.com/)\n"
                  << "Step 2: Export a Constant Contact CSV file of type 'Attendee report' for the event you are interested in, naming it in the format YYYY-MM-DD.TICKETTYPE.Registration Spreadsheet.csv. For example, 2026-03-24.New York City.Registration Spreadsheet.csv\n"
                  << "Step 3: Save the exported CSV file to the folder " << out_folder << "\n"
                  << "Step 4: Rerun " << g_exe_name << " with the CSV file you just exported.\n"
                  << "Step 5: Contact your Event Coordinator to report this error, including a copy of the file " << log_ref << "\n";
        return false;
    };

    // If auth failed and refresh couldn't help, trigger structured error
    if (auth_failed) {
        return api_failure("Authentication failed. Token refresh also failed. A new access token and refresh token are required.");
    }

    if (events_json.empty()) {
        return api_failure("Failed to fetch events list from Constant Contact API. Empty response.");
    }

    json events_resp;
    try {
        events_resp = json::parse(events_json);
    } catch (const json::parse_error& e) {
        return api_failure(std::string("Failed to parse events API response: ") + e.what()
                           + "\nResponse: " + events_json.substr(0, 500));
    }

    // Check for API error responses (e.g. 401, 403)
    // API may return errors as an array [{...}] or an object {...}
    if (events_resp.is_array() && !events_resp.empty()
        && events_resp[0].contains("error_key")) {
        return api_failure("API error: " + events_resp[0].value("error_key", "")
                           + " - " + events_resp[0].value("error_message", events_json.substr(0, 300)));
    }
    if (events_resp.is_object()
        && (events_resp.contains("error_key") || events_resp.contains("error_message"))) {
        return api_failure("API error: " + events_resp.value("error_key", "")
                           + " - " + events_resp.value("error_message", events_json.substr(0, 300)));
    }

    if (!events_resp.contains("records") || !events_resp["records"].is_array()
        || events_resp["records"].empty()) {
        // Retry with explicit event_status to include completed/ended events
        std::cout << "Retrying with explicit status filter...\n";
        events_json = http_get(
            L"api.cc.email",
            L"/v3/events?event_status=ACTIVE&limit=50",
            token, api_key);
        if (!events_json.empty()) {
            try { events_resp = json::parse(events_json); } catch (...) {}
        }
        if (!events_resp.contains("records") || !events_resp["records"].is_array()
            || events_resp["records"].empty()) {
            events_json = http_get(
                L"api.cc.email",
                L"/v3/events?event_status=COMPLETE&limit=50",
                token, api_key);
            if (!events_json.empty()) {
                try { events_resp = json::parse(events_json); } catch (...) {}
            }
        }

        if (!events_resp.contains("records") || !events_resp["records"].is_array()
            || events_resp["records"].empty()) {
            return api_failure("No events found after retries. Initial response: "
                               + events_json.substr(0, 500));
        }
    }

    auto& records = events_resp["records"];
    std::cout << "\nAvailable Events:\n";
    for (size_t i = 0; i < records.size(); ++i) {
        std::string name = records[i].value("name", "Unnamed Event");
        std::string status = records[i].value("status", "");
        std::cout << "  " << (i + 1) << ". " << name;
        if (!status.empty()) std::cout << " [" << status << "]";
        std::cout << "\n";
    }
    int ev_idx = 0;
    while (true) {
        std::cout << "\nChoose an event (enter number): ";
        std::string ev_choice;
        std::getline(std::cin, ev_choice);
        { std::string t = trim(ev_choice); ev_idx = (!t.empty() && std::all_of(t.begin(), t.end(), ::isdigit)) ? std::stoi(t) : 0; }
        if (ev_idx >= 1 && ev_idx <= (int)records.size()) break;
        std::cout << "Invalid Response\n";
    }

    std::string event_id = records[ev_idx - 1].value("event_id", "");
    if (event_id.empty()) {
        write_error_log(4, "Selected event has no ID.",
            "The selected event record from the Constant Contact API contains no event_id field.");
        std::cerr << "Error: Selected event has no ID.\n";
        return false;
    }

    // Build event title and extract date from event name
    {
        std::string name = records[ev_idx - 1].value("name", "Unnamed Event");
        out_event_title = name;
        out_event_date = parse_event_date(name);  // YYYY-MM-DD from "Month DD, YYYY ..." name
    }

    // Now that we have the event date, set the actual output CSV file path
    {
        std::string prefix = out_event_date.empty() ? "" : (out_event_date + ".");
        output_csv_path = output_dir / (prefix + "Registration Spreadsheet.csv");
    }

    // --- Get event details for the track_id ---
    std::cout << "Fetching event details...\n";
    std::wstring event_path = L"/v3/events/" + utf8_to_wide(event_id);
    std::string event_detail_json = http_get(L"api.cc.email", event_path, token, api_key);

    if (event_detail_json.empty()) {
        return api_failure("Failed to fetch event details. Empty response.");
    }

    json event_detail;
    try {
        event_detail = json::parse(event_detail_json);
    } catch (...) {
        return api_failure("Failed to parse event details response.");
    }

    std::string track_id;
    if (event_detail.contains("default_track")
        && event_detail["default_track"].contains("track_id")) {
        track_id = event_detail["default_track"]["track_id"].get<std::string>();
    }
    if (track_id.empty()) {
        return api_failure("Could not find track ID for this event.");
    }

    // --- Fetch registrations with pagination ---
    std::cout << "Fetching registrations...\n";
    std::vector<json> all_registrations;
    std::string next_cursor;

    do {
        std::wstring reg_path = L"/v3/events/" + utf8_to_wide(event_id)
            + L"/tracks/" + utf8_to_wide(track_id)
            + L"/registrations?registration_status=REGISTERED&limit=100";
        if (!next_cursor.empty()) {
            reg_path += L"&next=" + utf8_to_wide(next_cursor);
        }

        std::string reg_json = http_get(L"api.cc.email", reg_path, token, api_key);
        if (reg_json.empty()) break;

        json reg_resp;
        try {
            reg_resp = json::parse(reg_json);
        } catch (...) { break; }

        if (reg_resp.contains("records") && reg_resp["records"].is_array()) {
            for (auto& rec : reg_resp["records"]) {
                all_registrations.push_back(rec);
            }
        }

        next_cursor.clear();
        if (reg_resp.contains("next_cursor") && reg_resp["next_cursor"].is_string()) {
            next_cursor = reg_resp["next_cursor"].get<std::string>();
        }
    } while (!next_cursor.empty());

    std::cout << "  Retrieved " << all_registrations.size() << " registrations.\n";

    if (all_registrations.empty()) {
        return api_failure("No registrations found for this event.");
    }

    // --- Fetch detailed registration data for each registrant ---
    // The list endpoint returns flat data (no ticket names or company).
    // The individual GET endpoint returns full nested details.
    std::cout << "Fetching detailed registration data (this may take a moment)...\n";
    std::vector<json> detailed_registrations;
    for (size_t i = 0; i < all_registrations.size(); ++i) {
        std::string reg_id = all_registrations[i].value("registration_id", "");
        if (reg_id.empty()) {
            detailed_registrations.push_back(all_registrations[i]);
            continue;
        }

        std::wstring detail_path = L"/v3/events/" + utf8_to_wide(event_id)
            + L"/tracks/" + utf8_to_wide(track_id)
            + L"/registrations/" + utf8_to_wide(reg_id);

        std::string detail_json = http_get(L"api.cc.email", detail_path, token, api_key);
        if (!detail_json.empty()) {
            try {
                detailed_registrations.push_back(json::parse(detail_json));
            } catch (...) {
                detailed_registrations.push_back(all_registrations[i]);
            }
        } else {
            detailed_registrations.push_back(all_registrations[i]);
        }

        // Rate limiting: API allows 4 requests/sec
        if (i < all_registrations.size() - 1) {
            Sleep(260);
        }

        // Progress indicator
        if ((i + 1) % 10 == 0 || i + 1 == all_registrations.size()) {
            std::cout << "  " << (i + 1) << "/" << all_registrations.size() << "\n";
        }
    }

    // --- Build CSV from registrations ---
    // If the file is locked (e.g. open in Excel), warn the user
    if (fs::exists(output_csv_path)) {
        std::ofstream test(output_csv_path, std::ios::app);
        if (!test.is_open()) {
            std::cerr << "Error: CSV file is locked (close Excel first): "
                      << output_csv_path.string() << "\n";
            std::cout << "Close the file and press Enter to retry...";
            std::string dummy;
            std::getline(std::cin, dummy);
        }
    }
    std::ofstream csv_out(output_csv_path);
    if (!csv_out.is_open()) {
        write_error_log(5, "Cannot write Registration Spreadsheet CSV.",
            "Path: " + output_csv_path.string());
        std::cerr << "Error: Cannot create CSV file at " << output_csv_path.string() << "\n";
        return false;
    }

    // Header matching the expected format
    csv_out << "Ticket Type,Check-in Status,Registered By,Registered By Email,"
               "Attendee's Email Address,Attendee's First Name,Attendee's Last Name,"
               "Attendee's Company/Organization,Attendee's ISC2 Member ID,"
               "Optional Guest's Email Address,Optional Guest's First Name,"
               "Optional Guest's Last Name,Optional Guest's Company/Organization,"
               "Optional Guest's ISC2 Member ID\n";

    for (auto& reg : detailed_registrations) {
        // --- Helper: look up a field_value by field_label in an array of field objects ---
        auto field_val = [](const json& fields, const std::string& label) -> std::string {
            if (!fields.is_array()) return "";
            for (auto& f : fields) {
                if (f.value("field_label", "") == label)
                    return f.value("field_value", "");
            }
            return "";
        };
        // Case-insensitive partial match variant
        auto field_val_contains = [](const json& fields, const std::string& substr) -> std::string {
            if (!fields.is_array()) return "";
            std::string lower_sub = substr;
            for (auto& c : lower_sub) c = (char)std::tolower((unsigned char)c);
            for (auto& f : fields) {
                std::string label = f.value("field_label", "");
                std::string lower_label = label;
                for (auto& c : lower_label) c = (char)std::tolower((unsigned char)c);
                if (lower_label.find(lower_sub) != std::string::npos)
                    return f.value("field_value", "");
            }
            return "";
        };

        // --- Ticket type from tickets[0].ticket_name ---
        std::string ticket_type;
        if (reg.contains("tickets") && reg["tickets"].is_array()
            && !reg["tickets"].empty()) {
            ticket_type = reg["tickets"][0].value("ticket_name",
                reg["tickets"][0].value("name", ""));
        }
        if (ticket_type.empty()) ticket_type = "General Admission";

        // --- Check-in status ---
        std::string checkin_status = reg.value("checkin_status", "");

        // --- Contact fields (registered-by info): contact[] is array of {field_name, field_value} ---
        std::string reg_first, reg_last, reg_email;
        if (reg.contains("contact") && reg["contact"].is_array()) {
            auto& contact = reg["contact"];
            for (auto& f : contact) {
                std::string fn = f.value("field_name", "");
                std::string fv = f.value("field_value", "");
                if (fn == "CONTACT_FIRST_NAME") reg_first = fv;
                else if (fn == "CONTACT_LAST_NAME") reg_last = fv;
                else if (fn == "CONTACT_EMAIL_ADDRESS") reg_email = fv;
            }
        }
        // Flat-list fallback
        if (reg_first.empty()) reg_first = reg.value("first_name", "");
        if (reg_last.empty())  reg_last  = reg.value("last_name", "");
        if (reg_email.empty()) reg_email = reg.value("email_address", "");

        std::string reg_by = reg_first;
        if (!reg_last.empty()) {
            if (!reg_by.empty()) reg_by += " ";
            reg_by += reg_last;
        }

        // --- Attendee fields from tickets[0].fields[] ---
        std::string att_email, att_first, att_last, att_company, att_member_id;
        std::string guest_email, guest_first, guest_last, guest_company, guest_member_id;

        if (reg.contains("tickets") && reg["tickets"].is_array()
            && !reg["tickets"].empty()
            && reg["tickets"][0].contains("fields")) {
            auto& tf = reg["tickets"][0]["fields"];
            att_email   = field_val(tf, "Attendee's Email Address");
            att_first   = field_val(tf, "Attendee's First Name");
            att_last    = field_val(tf, "Attendee's Last Name");
            att_company = field_val(tf, "Attendee's Company/Organization");
            att_member_id = field_val_contains(tf, "isc2");
            if (att_member_id.empty())
                att_member_id = field_val_contains(tf, "member id");

            guest_email   = field_val_contains(tf, "guest's email");
            if (guest_email.empty())
                guest_email = field_val(tf, "Optional Guest's Email Address");
            guest_first   = field_val_contains(tf, "guest's first");
            if (guest_first.empty())
                guest_first = field_val(tf, "Optional Guest's First Name");
            guest_last    = field_val_contains(tf, "guest's last");
            if (guest_last.empty())
                guest_last = field_val(tf, "Optional Guest's Last Name");
            guest_company = field_val_contains(tf, "guest's company");
            if (guest_company.empty())
                guest_company = field_val(tf, "Optional Guest's Company/Organization");
            guest_member_id = field_val_contains(tf, "guest's isc2");
            if (guest_member_id.empty())
                guest_member_id = field_val_contains(tf, "guest's member");
        }

        // If attendee fields weren't in ticket, fall back to contact info
        if (att_first.empty()) att_first = reg_first;
        if (att_last.empty())  att_last  = reg_last;
        if (att_email.empty()) att_email = reg_email;

        csv_out << csv_escape(ticket_type) << ","       // 0  Ticket Type
                << csv_escape(checkin_status) << ","    // 1  Check-in Status
                << csv_escape(reg_by) << ","            // 2  Registered By
                << csv_escape(reg_email) << ","         // 3  Registered By Email
                << csv_escape(att_email) << ","         // 4  Attendee Email
                << csv_escape(att_first) << ","         // 5  Attendee First
                << csv_escape(att_last) << ","          // 6  Attendee Last
                << csv_escape(att_company) << ","       // 7  Company
                << csv_escape(att_member_id) << ","     // 8  Member ID
                << csv_escape(guest_email) << ","       // 9  Guest Email
                << csv_escape(guest_first) << ","       // 10 Guest First
                << csv_escape(guest_last) << ","        // 11 Guest Last
                << csv_escape(guest_company) << ","     // 12 Guest Company
                << csv_escape(guest_member_id)          // 13 Guest Member ID
                << "\n";
    }

    csv_out.close();
    std::cout << "Registration data saved to: " << output_csv_path.string() << "\n";
    return true;
}

// ============================================================================
// DOCX ZIP manipulation using Windows Shell API
// ============================================================================

static bool extract_zip(const fs::path& zip_path, const fs::path& dest_dir) {
    fs::create_directories(dest_dir);

    // Use .NET ZipFile class via PowerShell — works in both PS 5.1 and 7
    std::string cmd = "powershell -NoProfile -Command \""
        "Add-Type -AssemblyName System.IO.Compression.FileSystem; "
        "[System.IO.Compression.ZipFile]::ExtractToDirectory('"
        + zip_path.string() + "', '" + dest_dir.string() + "')\"";
    return system(cmd.c_str()) == 0;
}

static bool create_zip(const fs::path& source_dir, const fs::path& zip_path) {
    // Remove existing zip
    if (fs::exists(zip_path)) fs::remove(zip_path);

    // Use .NET ZipFile class via PowerShell — works in both PS 5.1 and 7
    std::string cmd = "powershell -NoProfile -Command \""
        "Add-Type -AssemblyName System.IO.Compression.FileSystem; "
        "[System.IO.Compression.ZipFile]::CreateFromDirectory('"
        + source_dir.string() + "', '" + zip_path.string() + "')\"";
    return system(cmd.c_str()) == 0;
}

// ============================================================================
// DOCX manipulation: Attendance Sheet
// ============================================================================

static bool populate_attendance_sheet(const fs::path& docx_path,
                                      const std::vector<std::string>& names,
                                      const std::string& event_title = "",
                                      const std::string& venue_name = "") {
    try {
    // Extract DOCX
    fs::path temp_dir = fs::temp_directory_path() / "isc2_attendance";
    if (fs::exists(temp_dir)) fs::remove_all(temp_dir);

    fs::path zip_copy = fs::temp_directory_path() / "isc2_attendance.zip";
    if (fs::exists(zip_copy)) fs::remove(zip_copy);
    fs::copy_file(docx_path, zip_copy);

    if (!extract_zip(zip_copy, temp_dir)) {
        std::cerr << "Error: Failed to extract Attendance Sheet DOCX.\n";
        return false;
    }

    // Parse document.xml
    fs::path doc_xml = temp_dir / "word" / "document.xml";
    pugi::xml_document doc;
    pugi::xml_parse_result parse_result = doc.load_file(doc_xml.c_str());
    if (!parse_result) {
        std::cerr << "Error: Failed to parse document.xml: " << parse_result.description() << "\n";
        return false;
    }

    // Helper: replace text that may span multiple w:r runs in a paragraph
    auto replace_across_runs = [](pugi::xml_node body_node,
                                   const std::string& find_text,
                                   const std::string& replace_text) {
        size_t find_len = find_text.size();
        for (auto p = body_node.child("w:p"); p; p = p.next_sibling("w:p")) {
            struct RunText { pugi::xml_node t_node; std::string text; };
            std::vector<RunText> runs;
            for (auto r = p.child("w:r"); r; r = r.next_sibling("w:r")) {
                auto t = r.child("w:t");
                if (t) runs.push_back({t, std::string(t.child_value())});
            }
            std::string para_text;
            for (auto& rt : runs) para_text += rt.text;

            size_t pos = para_text.find(find_text);
            if (pos == std::string::npos) continue;

            size_t char_offset = 0;
            bool replaced = false;
            for (size_t i = 0; i < runs.size(); ++i) {
                size_t run_start = char_offset;
                size_t run_end = char_offset + runs[i].text.size();
                if (!replaced && run_start <= pos && run_end > pos) {
                    std::string before = runs[i].text.substr(0, pos - run_start);
                    size_t match_here = std::min(find_len, run_end - pos);
                    std::string after_match;
                    if (match_here >= find_len)
                        after_match = runs[i].text.substr(pos - run_start + find_len);
                    runs[i].t_node.text().set((before + replace_text + after_match).c_str());
                    runs[i].t_node.attribute("xml:space").set_value("preserve");
                    replaced = true;
                    size_t remaining = find_len - match_here;
                    for (size_t j = i + 1; j < runs.size() && remaining > 0; ++j) {
                        size_t consume = std::min(remaining, runs[j].text.size());
                        runs[j].t_node.text().set(runs[j].text.substr(consume).c_str());
                        remaining -= consume;
                    }
                    break;
                }
                char_offset = run_end;
            }
            if (replaced) break;
        }
    };

    // Replace "Event Title" placeholder with the actual event title
    if (!event_title.empty()) {
        auto body_node = doc.child("w:document").child("w:body");
        replace_across_runs(body_node, "Event Title", event_title);
    }

    // Replace "New York City" within "Attendance Sheet – New York City Venue"
    // Only target the paragraph that contains both "Attendance Sheet" and "Venue"
    if (!venue_name.empty()) {
        auto body_node = doc.child("w:document").child("w:body");
        for (auto p = body_node.child("w:p"); p; p = p.next_sibling("w:p")) {
            // Concatenate all run text for this paragraph
            std::string para_text;
            for (auto r = p.child("w:r"); r; r = r.next_sibling("w:r")) {
                auto t = r.child("w:t");
                if (t) para_text += t.child_value();
            }
            // Only replace in the paragraph containing "Attendance Sheet" and "Venue"
            if (para_text.find("Attendance Sheet") != std::string::npos
                && para_text.find("Venue") != std::string::npos
                && para_text.find("New York City") != std::string::npos) {
                // Use replace_across_runs scoped to just this paragraph's parent
                // but since the helper scans from body, we do inline replacement here
                struct RunText { pugi::xml_node t_node; std::string text; };
                std::vector<RunText> runs;
                for (auto r = p.child("w:r"); r; r = r.next_sibling("w:r")) {
                    auto t = r.child("w:t");
                    if (t) runs.push_back({t, std::string(t.child_value())});
                }
                size_t pos = para_text.find("New York City");
                size_t find_len = 13; // strlen("New York City")
                size_t char_offset = 0;
                for (size_t i = 0; i < runs.size(); ++i) {
                    size_t run_start = char_offset;
                    size_t run_end = char_offset + runs[i].text.size();
                    if (run_start <= pos && run_end > pos) {
                        std::string before = runs[i].text.substr(0, pos - run_start);
                        size_t match_here = std::min(find_len, run_end - pos);
                        std::string after_match;
                        if (match_here >= find_len)
                            after_match = runs[i].text.substr(pos - run_start + find_len);
                        runs[i].t_node.text().set((before + venue_name + after_match).c_str());
                        runs[i].t_node.attribute("xml:space").set_value("preserve");
                        size_t remaining = find_len - match_here;
                        for (size_t j = i + 1; j < runs.size() && remaining > 0; ++j) {
                            size_t consume = std::min(remaining, runs[j].text.size());
                            runs[j].t_node.text().set(runs[j].text.substr(consume).c_str());
                            remaining -= consume;
                        }
                        break;
                    }
                    char_offset = run_end;
                }
                break; // Only replace in this one paragraph
            }
        }
    }

    // Find the first table (attendance sign-in table with Name/Signature columns)
    // The table has 5 columns: Name | Signature | (spacer) | Name | Signature
    // Row 0 is the header row. We need to fill data rows.
    auto body = doc.child("w:document").child("w:body");
    pugi::xml_node attendance_table;

    for (auto tbl = body.child("w:tbl"); tbl; tbl = tbl.next_sibling("w:tbl")) {
        // Check if this table has "Name" and "Signature" headers
        auto first_row = tbl.child("w:tr");
        if (!first_row) continue;

        bool has_name = false, has_sig = false;
        for (auto tc = first_row.child("w:tc"); tc; tc = tc.next_sibling("w:tc")) {
            for (auto p = tc.child("w:p"); p; p = p.next_sibling("w:p")) {
                for (auto r = p.child("w:r"); r; r = r.next_sibling("w:r")) {
                    auto t = r.child("w:t");
                    if (t) {
                        std::string text = t.child_value();
                        if (text == "Name") has_name = true;
                        if (text == "Signature") has_sig = true;
                    }
                }
            }
        }
        if (has_name && has_sig) {
            attendance_table = tbl;
            break;
        }
    }

    if (!attendance_table) {
        std::cerr << "Error: Could not find attendance table in document.\n";
        return false;
    }

    // Get data rows (skip header row 0)
    std::vector<pugi::xml_node> data_rows;
    bool first = true;
    for (auto tr = attendance_table.child("w:tr"); tr; tr = tr.next_sibling("w:tr")) {
        if (first) { first = false; continue; }
        data_rows.push_back(tr);
    }

    // Template has 5 columns per row: col0=Left Name, col1=Left Signature, col2=Spacer, col3=Right Name, col4=Right Signature
    // Fill left Name column first (col0), then right Name column (col3)
    size_t rows_available = data_rows.size();
    size_t left_capacity = rows_available;   // one name column on the left
    size_t right_capacity = rows_available;  // one name column on the right
    size_t total_capacity = left_capacity + right_capacity;

    if (names.size() > total_capacity) {
        std::cout << "Note: " << names.size() << " names span "
                  << ((names.size() + total_capacity - 1) / total_capacity)
                  << " pages.\n";
    }

    // Helper to find/create a text run in a table cell
    auto set_cell_text = [](pugi::xml_node tc, const std::string& text) {
        // Find or create w:p -> w:r -> w:t
        auto p = tc.child("w:p");
        if (!p) p = tc.append_child("w:p");

        auto r = p.child("w:r");
        if (!r) {
            r = p.append_child("w:r");
            // Copy run properties from a sibling if possible
            auto pPr = p.child("w:pPr");
            if (pPr) {
                auto rPr_ref = pPr.child("w:rPr");
                if (rPr_ref) {
                    auto rPr = r.prepend_child("w:rPr");
                    for (auto child = rPr_ref.first_child(); child; child = child.next_sibling()) {
                        rPr.append_copy(child);
                    }
                }
            }
        }

        auto t = r.child("w:t");
        if (!t) t = r.append_child("w:t");
        t.text().set(text.c_str());
        // Set xml:space="preserve" only if not already present — appending a
        // duplicate attribute produces invalid XML that Word refuses to open.
        if (!t.attribute("xml:space"))
            t.append_attribute("xml:space").set_value("preserve");
    };

    // Fill names: left column first, then right
    for (size_t i = 0; i < names.size() && i < total_capacity; ++i) {
        size_t row_idx;
        int col_idx;

        if (i < left_capacity) {
            row_idx = i;
            col_idx = 0;  // Left Name column
        } else {
            row_idx = i - left_capacity;
            col_idx = 3;  // Right Name column
        }

        auto tr = data_rows[row_idx];

        // Get the target cell (column index)
        pugi::xml_node tc;
        int c = 0;
        for (auto cell = tr.child("w:tc"); cell; cell = cell.next_sibling("w:tc"), ++c) {
            if (c == col_idx) {
                tc = cell;
                break;
            }
        }

        if (tc) {
            set_cell_text(tc, names[i]);
        }
    }

    // If names exceed the first page, add additional pages by cloning the entire table
    if (names.size() > total_capacity) {
        size_t overflow_start = total_capacity;
        pugi::xml_node last_node = attendance_table;

        while (overflow_start < names.size()) {
            // Insert a page break paragraph after the last inserted node
            auto pb_para = body.insert_child_after("w:p", last_node);
            auto pb_r = pb_para.append_child("w:r");
            auto pb_br = pb_r.append_child("w:br");
            pb_br.append_attribute("w:type").set_value("page");
            last_node = pb_para;

            // Clone the original attendance table and insert it after the page break
            auto new_table = body.insert_copy_after(attendance_table, last_node);
            last_node = new_table;

            // Collect data rows of new table (skip header) and clear all text
            std::vector<pugi::xml_node> new_data_rows;
            bool first_row2 = true;
            for (auto tr = new_table.child("w:tr"); tr; tr = tr.next_sibling("w:tr")) {
                if (first_row2) { first_row2 = false; continue; }
                new_data_rows.push_back(tr);
                for (auto tc = tr.child("w:tc"); tc; tc = tc.next_sibling("w:tc")) {
                    for (auto p = tc.child("w:p"); p; p = p.next_sibling("w:p")) {
                        for (auto r = p.child("w:r"); r; r = r.next_sibling("w:r")) {
                            auto t = r.child("w:t");
                            if (t) t.text().set("");
                        }
                    }
                }
            }

            // Fill new page: left column first, then right
            size_t page_left  = new_data_rows.size();
            size_t page_total = page_left * 2;
            for (size_t i = 0; i < page_total && overflow_start + i < names.size(); ++i) {
                size_t row_idx = (i < page_left) ? i : (i - page_left);
                int    col_idx = (i < page_left) ? 0 : 3;

                auto tr = new_data_rows[row_idx];
                pugi::xml_node tc2;
                int c = 0;
                for (auto cell = tr.child("w:tc"); cell; cell = cell.next_sibling("w:tc"), ++c) {
                    if (c == col_idx) { tc2 = cell; break; }
                }
                if (tc2) set_cell_text(tc2, names[overflow_start + i]);
            }

            overflow_start += page_total;
        }
    }

    // Make gray (#595959) run-level font colors pure black and remove yellow
    // highlighting from all runs in document.xml.
    {
        struct ColorFixer {
            static void fix(pugi::xml_node node) {
                for (auto child = node.first_child(); child; ) {
                    auto next = child.next_sibling();
                    std::string cname = child.name();
                    if (cname == "w:color") {
                        auto val_attr = child.attribute("w:val");
                        if (val_attr && std::string(val_attr.as_string()) == "595959") {
                            val_attr.set_value("000000");
                            child.remove_attribute("w:themeColor");
                            child.remove_attribute("w:themeTint");
                            child.remove_attribute("w:themeShade");
                        }
                    } else if (cname == "w:highlight") {
                        // Remove yellow (or any) highlight element entirely
                        node.remove_child(child);
                        child = next;
                        continue;
                    }
                    fix(child);
                    child = next;
                }
            }
        };
        ColorFixer::fix(doc.child("w:document"));
    }

    // Minimize inter-table paragraphs to prevent blank pages between tables.
    // Word requires a w:p between two w:tbl nodes but it can be made 1pt/0-spacing
    // so it never overflows onto its own page.
    {
        pugi::xml_node prev_tbl;
        for (auto node = body.first_child(); node; node = node.next_sibling()) {
            std::string name = node.name();
            if (name == "w:tbl") {
                if (prev_tbl) {
                    // Minimize every w:p that sits between prev_tbl and this w:tbl
                    for (auto sib = prev_tbl.next_sibling(); sib && sib != node;
                         sib = sib.next_sibling()) {
                        if (std::string(sib.name()) != "w:p") continue;
                        // Ensure w:pPr exists
                        auto pPr = sib.child("w:pPr");
                        if (!pPr) pPr = sib.prepend_child("w:pPr");
                        // Zero spacing before/after
                        auto spacing = pPr.child("w:spacing");
                        if (!spacing) spacing = pPr.append_child("w:spacing");
                        if (!spacing.attribute("w:before")) spacing.append_attribute("w:before").set_value("0");
                        else spacing.attribute("w:before").set_value("0");
                        if (!spacing.attribute("w:after")) spacing.append_attribute("w:after").set_value("0");
                        else spacing.attribute("w:after").set_value("0");
                        // Paragraph mark run properties: 1pt font (sz=2)
                        auto rPr = pPr.child("w:rPr");
                        if (!rPr) rPr = pPr.append_child("w:rPr");
                        auto sz = rPr.child("w:sz");
                        if (!sz) sz = rPr.append_child("w:sz");
                        if (!sz.attribute("w:val")) sz.append_attribute("w:val").set_value("2");
                        else sz.attribute("w:val").set_value("2");
                        auto szCs = rPr.child("w:szCs");
                        if (!szCs) szCs = rPr.append_child("w:szCs");
                        if (!szCs.attribute("w:val")) szCs.append_attribute("w:val").set_value("2");
                        else szCs.attribute("w:val").set_value("2");
                    }
                }
                prev_tbl = node;
            }
        }
    }

    // Save the modified XML
    if (!doc.save_file(doc_xml.c_str(), "  ", pugi::format_raw | pugi::format_no_declaration)) {
        std::cerr << "Error: Failed to save modified document.xml.\n";
        return false;
    }

    // Also need to preserve the XML declaration
    {
        std::ifstream in(doc_xml);
        std::string content((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
        in.close();

        if (content.find("<?xml") == std::string::npos) {
            std::ofstream out(doc_xml);
            out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
            out << content;
        }
    }

    // Repackage as DOCX
    if (fs::exists(docx_path)) fs::remove(docx_path);

    fs::path new_zip = fs::temp_directory_path() / "isc2_attendance_new.zip";
    if (!create_zip(temp_dir, new_zip)) {
        std::cerr << "Error: Failed to re-create DOCX zip.\n";
        return false;
    }

    fs::copy_file(new_zip, docx_path, fs::copy_options::overwrite_existing);

    // Cleanup
    fs::remove_all(temp_dir);
    fs::remove(zip_copy);
    fs::remove(new_zip);

    return true;
    } catch (const std::exception& e) {
        std::cerr << "Error: Exception in populate_attendance_sheet: " << e.what() << "\n";
        return false;
    }
}

// ============================================================================
// DOCX manipulation: Name Tag Template
// ============================================================================

static bool populate_name_tags(const fs::path& docx_path,
                               const std::vector<std::string>& names) {
    try {
    // Extract DOCX
    fs::path temp_dir = fs::temp_directory_path() / "isc2_nametag";
    if (fs::exists(temp_dir)) fs::remove_all(temp_dir);

    fs::path zip_copy = fs::temp_directory_path() / "isc2_nametag.zip";
    if (fs::exists(zip_copy)) fs::remove(zip_copy);
    fs::copy_file(docx_path, zip_copy);

    if (!extract_zip(zip_copy, temp_dir)) {
        std::cerr << "Error: Failed to extract Name Tag DOCX.\n";
        return false;
    }

    // Parse document.xml
    fs::path doc_xml = temp_dir / "word" / "document.xml";
    pugi::xml_document doc;
    pugi::xml_parse_result parse_result = doc.load_file(doc_xml.c_str());
    if (!parse_result) {
        std::cerr << "Error: Failed to parse document.xml: " << parse_result.description() << "\n";
        return false;
    }

    // Find all occurrences of "Name" text in <w:t> elements
    struct NameNode {
        pugi::xml_node text_node;
        pugi::xml_node table_node;  // parent table for page tracking
    };

    std::vector<NameNode> name_nodes;
    auto body = doc.child("w:document").child("w:body");

    // Walk through tables to find "Name" text nodes
    // Note: some cells have " Name" (space-prefixed with xml:space="preserve")
    // so we must trim whitespace before comparing.
    for (auto tbl = body.child("w:tbl"); tbl; tbl = tbl.next_sibling("w:tbl")) {
        // Walk all w:t elements within this table
        for (auto tr = tbl.child("w:tr"); tr; tr = tr.next_sibling("w:tr")) {
            for (auto tc = tr.child("w:tc"); tc; tc = tc.next_sibling("w:tc")) {
                for (auto p = tc.child("w:p"); p; p = p.next_sibling("w:p")) {
                    for (auto r = p.child("w:r"); r; r = r.next_sibling("w:r")) {
                        auto t = r.child("w:t");
                        if (t && trim(std::string(t.child_value())) == "Name") {
                            name_nodes.push_back({t, tbl});
                        }
                    }
                }
            }
        }
    }

    std::cout << "  Found " << name_nodes.size() << " 'Name' placeholders in Name Tag template.\n";

    if (names.size() > name_nodes.size()) {
        std::cerr << "  ERROR: Number of attendees (" << names.size()
                  << ") exceeds available name tag slots (" << name_nodes.size() << ")!\n";
    }

    // Replace "Name" with actual names; clear remaining placeholders on partial pages
    size_t replaced = 0;
    for (size_t i = 0; i < name_nodes.size(); ++i) {
        if (i < names.size()) {
            name_nodes[i].text_node.text().set(names[i].c_str());
            ++replaced;
        } else {
            // Clear unused placeholder so it prints blank
            name_nodes[i].text_node.text().set("");
        }
    }

    std::cout << "  Replaced " << replaced << " name tag placeholders.\n";

    // Track which tables had at least one replacement
    std::set<pugi::xml_node> tables_with_replacements;
    for (size_t i = 0; i < replaced; ++i) {
        tables_with_replacements.insert(name_nodes[i].table_node);
    }

    // Collect all tables that contain "Name" placeholders (each = one page of labels)
    std::vector<pugi::xml_node> all_name_tables;
    {
        std::set<pugi::xml_node> seen;
        for (auto& nn : name_nodes) {
            if (seen.find(nn.table_node) == seen.end()) {
                seen.insert(nn.table_node);
                all_name_tables.push_back(nn.table_node);
            }
        }
    }

    // Remove pages (tables) where NO "Name" placeholder was replaced.
    // Each "page" in the name tag document consists of a table (holding the
    // Name text) PLUS several surrounding paragraphs that contain the label
    // outline shapes (rounded rectangles).  We must remove the table AND all
    // associated paragraphs to fully eliminate the empty page.
    std::set<pugi::xml_node> tables_to_remove;
    for (auto& tbl : all_name_tables) {
        if (tables_with_replacements.find(tbl) == tables_with_replacements.end()) {
            tables_to_remove.insert(tbl);
        }
    }

    size_t pages_removed = 0;
    for (auto& tbl : all_name_tables) {
        if (tables_to_remove.find(tbl) == tables_to_remove.end()) continue;

        // Remove all preceding w:p siblings until we hit another w:tbl or
        // the beginning of the body.
        std::vector<pugi::xml_node> preceding_paragraphs;
        for (auto prev = tbl.previous_sibling(); prev; prev = prev.previous_sibling()) {
            if (std::string(prev.name()) == "w:p") {
                preceding_paragraphs.push_back(prev);
            } else {
                break;  // hit a table or other element — stop
            }
        }
        for (auto& p : preceding_paragraphs) {
            body.remove_child(p);
        }

        body.remove_child(tbl);
        ++pages_removed;
    }

    // Final sweep: remove any orphaned paragraphs (label outlines) that
    // remain between the last kept table and the <w:sectPr> element.
    // These are leftovers from removed pages whose preceding paragraphs
    // were stranded because the removal order didn't catch them.
    {
        auto sect = body.child("w:sectPr");
        if (sect) {
            std::vector<pugi::xml_node> orphans;
            for (auto prev = sect.previous_sibling(); prev; prev = prev.previous_sibling()) {
                if (std::string(prev.name()) == "w:p") {
                    orphans.push_back(prev);
                } else {
                    break;  // hit the last kept table — stop
                }
            }
            for (auto& p : orphans) {
                body.remove_child(p);
            }
        }
    }

    if (pages_removed > 0) {
        std::cout << "  Removed " << pages_removed << " unused name tag pages.\n";
    }

    // Save the modified XML
    if (!doc.save_file(doc_xml.c_str(), "  ", pugi::format_raw | pugi::format_no_declaration)) {
        std::cerr << "Error: Failed to save modified document.xml.\n";
        return false;
    }

    // Preserve XML declaration
    {
        std::ifstream in(doc_xml);
        std::string content((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
        in.close();

        if (content.find("<?xml") == std::string::npos) {
            std::ofstream out(doc_xml);
            out << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
            out << content;
        }
    }

    // Repackage as DOCX
    if (fs::exists(docx_path)) fs::remove(docx_path);

    fs::path new_zip = fs::temp_directory_path() / "isc2_nametag_new.zip";
    if (!create_zip(temp_dir, new_zip)) {
        std::cerr << "Error: Failed to re-create DOCX zip.\n";
        return false;
    }

    fs::copy_file(new_zip, docx_path, fs::copy_options::overwrite_existing);

    // Cleanup
    fs::remove_all(temp_dir);
    fs::remove(zip_copy);
    fs::remove(new_zip);

    return true;
    } catch (const std::exception& e) {
        std::cerr << "Error: Exception in populate_name_tags: " << e.what() << "\n";
        return false;
    }
}

// ============================================================================
// PDF Conversion using Word COM Automation
// ============================================================================

static bool convert_to_pdf(const fs::path& doc_path, const fs::path& pdf_path,
                          bool hide_last_paragraph = false) {
    std::cout << "  Converting " << doc_path.filename().string() << " to PDF...\n";

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    bool we_initialized = SUCCEEDED(hr);

    // Use Word COM automation
    CLSID clsid;
    hr = CLSIDFromProgID(L"Word.Application", &clsid);
    if (FAILED(hr)) {
        std::cerr << "Error: Microsoft Word is not installed or not registered.\n";
        std::cerr << "  Word is required for PDF conversion.\n";
        if (we_initialized) CoUninitialize();
        return false;
    }

    IDispatch* pWordApp = nullptr;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER,
                          IID_IDispatch, (void**)&pWordApp);
    if (FAILED(hr) || !pWordApp) {
        std::cerr << "Error: Failed to start Microsoft Word.\n";
        if (we_initialized) CoUninitialize();
        return false;
    }

    // Helper lambda for IDispatch calls
    auto invoke = [](IDispatch* pDisp, const wchar_t* name, WORD flags,
                     VARIANT* result, int argc = 0, VARIANT* args = nullptr) -> HRESULT {
        DISPID dispid;
        LPOLESTR name_ptr = const_cast<LPOLESTR>(name);
        HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &name_ptr, 1, LOCALE_USER_DEFAULT, &dispid);
        if (FAILED(hr)) return hr;

        DISPPARAMS dp = {};
        dp.cArgs = argc;
        dp.rgvarg = args;
        if (flags & DISPATCH_PROPERTYPUT) {
            DISPID putid = DISPID_PROPERTYPUT;
            dp.cNamedArgs = 1;
            dp.rgdispidNamedArgs = &putid;
        }

        return pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, flags, &dp, result, nullptr, nullptr);
    };

    bool success = false;

    // Set Word.Visible = False
    {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_BOOL;
        val.boolVal = VARIANT_FALSE;
        invoke(pWordApp, L"Visible", DISPATCH_PROPERTYPUT, nullptr, 1, &val);
    }

    // Set DisplayAlerts = False (wdAlertsNone = 0)
    {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = 0; // wdAlertsNone
        invoke(pWordApp, L"DisplayAlerts", DISPATCH_PROPERTYPUT, nullptr, 1, &val);
    }

    // Get Documents collection
    VARIANT docs_result;
    VariantInit(&docs_result);
    hr = invoke(pWordApp, L"Documents", DISPATCH_PROPERTYGET, &docs_result);

    if (SUCCEEDED(hr) && docs_result.vt == VT_DISPATCH && docs_result.pdispVal) {
        IDispatch* pDocs = docs_result.pdispVal;

        // Open the document
        std::wstring abs_doc_path = fs::absolute(doc_path).wstring();
        VARIANT open_args[3];
        VariantInit(&open_args[0]);
        VariantInit(&open_args[1]);
        VariantInit(&open_args[2]);

        // Args are in reverse order for IDispatch
        open_args[2].vt = VT_BSTR;
        open_args[2].bstrVal = SysAllocString(abs_doc_path.c_str());
        open_args[1].vt = VT_BOOL;
        open_args[1].boolVal = VARIANT_FALSE;  // ConfirmConversions
        open_args[0].vt = VT_BOOL;
        open_args[0].boolVal = hide_last_paragraph ? VARIANT_FALSE : VARIANT_TRUE;   // ReadOnly

        VARIANT doc_result;
        VariantInit(&doc_result);
        hr = invoke(pDocs, L"Open", DISPATCH_METHOD, &doc_result, 3, open_args);

        SysFreeString(open_args[2].bstrVal);

        if (SUCCEEDED(hr) && doc_result.vt == VT_DISPATCH && doc_result.pdispVal) {
            IDispatch* pDoc = doc_result.pdispVal;

            // Hide the trailing paragraph to prevent a blank last page.
            // Word always inserts a mandatory paragraph after a table; when the
            // table fills the page it overflows to a blank page.  Making it
            // hidden removes it from the printed/PDF output.
            if (hide_last_paragraph) {
                VARIANT paras_result;
                VariantInit(&paras_result);
                hr = invoke(pDoc, L"Paragraphs", DISPATCH_PROPERTYGET, &paras_result);
                if (SUCCEEDED(hr) && paras_result.vt == VT_DISPATCH && paras_result.pdispVal) {
                    IDispatch* pParas = paras_result.pdispVal;
                    VARIANT last_para_result;
                    VariantInit(&last_para_result);
                    hr = invoke(pParas, L"Last", DISPATCH_PROPERTYGET, &last_para_result);
                    if (SUCCEEDED(hr) && last_para_result.vt == VT_DISPATCH && last_para_result.pdispVal) {
                        IDispatch* pLastPara = last_para_result.pdispVal;
                        VARIANT range_result;
                        VariantInit(&range_result);
                        hr = invoke(pLastPara, L"Range", DISPATCH_PROPERTYGET, &range_result);
                        if (SUCCEEDED(hr) && range_result.vt == VT_DISPATCH && range_result.pdispVal) {
                            IDispatch* pRange = range_result.pdispVal;

                            // Set Font.Hidden = True, Size = 1
                            VARIANT font_result;
                            VariantInit(&font_result);
                            hr = invoke(pRange, L"Font", DISPATCH_PROPERTYGET, &font_result);
                            if (SUCCEEDED(hr) && font_result.vt == VT_DISPATCH && font_result.pdispVal) {
                                IDispatch* pFont = font_result.pdispVal;
                                VARIANT hv; VariantInit(&hv); hv.vt = VT_BOOL; hv.boolVal = VARIANT_TRUE;
                                invoke(pFont, L"Hidden", DISPATCH_PROPERTYPUT, nullptr, 1, &hv);
                                VARIANT sv; VariantInit(&sv); sv.vt = VT_R4; sv.fltVal = 1.0f;
                                invoke(pFont, L"Size", DISPATCH_PROPERTYPUT, nullptr, 1, &sv);
                                pFont->Release();
                            }

                            // Set paragraph spacing to zero, line spacing exact 1pt
                            VARIANT pf_result;
                            VariantInit(&pf_result);
                            hr = invoke(pRange, L"ParagraphFormat", DISPATCH_PROPERTYGET, &pf_result);
                            if (SUCCEEDED(hr) && pf_result.vt == VT_DISPATCH && pf_result.pdispVal) {
                                IDispatch* pPF = pf_result.pdispVal;
                                VARIANT zv; VariantInit(&zv); zv.vt = VT_R4; zv.fltVal = 0.0f;
                                invoke(pPF, L"SpaceBefore", DISPATCH_PROPERTYPUT, nullptr, 1, &zv);
                                invoke(pPF, L"SpaceAfter", DISPATCH_PROPERTYPUT, nullptr, 1, &zv);
                                VARIANT lr; VariantInit(&lr); lr.vt = VT_I4; lr.lVal = 4;
                                invoke(pPF, L"LineSpacingRule", DISPATCH_PROPERTYPUT, nullptr, 1, &lr);
                                VARIANT ls; VariantInit(&ls); ls.vt = VT_R4; ls.fltVal = 1.0f;
                                invoke(pPF, L"LineSpacing", DISPATCH_PROPERTYPUT, nullptr, 1, &ls);
                                pPF->Release();
                            }
                            pRange->Release();
                        }
                        pLastPara->Release();
                    }
                    pParas->Release();
                }
                std::cout << "  Hidden trailing paragraph to eliminate blank page.\n";
            }

            // ExportAsFixedFormat - save as PDF
            // ExportAsFixedFormat(OutputFileName, ExportFormat, ...)
            // wdExportFormatPDF = 17
            std::wstring abs_pdf_path = fs::absolute(pdf_path).wstring();

            VARIANT export_args[2];
            VariantInit(&export_args[0]);
            VariantInit(&export_args[1]);

            export_args[1].vt = VT_BSTR;
            export_args[1].bstrVal = SysAllocString(abs_pdf_path.c_str());
            export_args[0].vt = VT_I4;
            export_args[0].lVal = 17;  // wdExportFormatPDF

            hr = invoke(pDoc, L"ExportAsFixedFormat", DISPATCH_METHOD, nullptr, 2, export_args);

            SysFreeString(export_args[1].bstrVal);

            if (SUCCEEDED(hr)) {
                success = true;
                std::cout << "  PDF created: " << pdf_path.string() << "\n";
            } else {
                std::cerr << "  Error: ExportAsFixedFormat failed (HRESULT: 0x"
                          << std::hex << hr << std::dec << ")\n";
            }

            // Close document — save if we modified it
            VARIANT close_arg;
            VariantInit(&close_arg);
            close_arg.vt = VT_I4;
            close_arg.lVal = hide_last_paragraph ? -1 : 0;  // wdSaveChanges or wdDoNotSaveChanges
            invoke(pDoc, L"Close", DISPATCH_METHOD, nullptr, 1, &close_arg);

            pDoc->Release();
        } else {
            std::cerr << "  Error: Failed to open document in Word.\n";
        }

        pDocs->Release();
    }

    // Quit Word
    VARIANT quit_arg;
    VariantInit(&quit_arg);
    quit_arg.vt = VT_I4;
    quit_arg.lVal = 0;  // wdDoNotSaveChanges
    invoke(pWordApp, L"Quit", DISPATCH_METHOD, nullptr, 1, &quit_arg);

    pWordApp->Release();
    if (we_initialized) CoUninitialize();

    return success;
}

// ============================================================================
// Main Program
// ============================================================================

static int run(int argc, char* argv[]) {
    std::cout << "=== ISC2 Event Coordinator ===\n";
    std::cout << "Generates Attendance Sheet and Name Tag documents from Constant Contact data.\n\n";

    // Determine base directory (where the executable is, or the project root)
    fs::path exe_path = fs::path(argv[0]).parent_path();
    fs::path base_dir;

    // Look for input templates folder relative to exe
    if (fs::exists(exe_path / "input templates")) {
        base_dir = exe_path;
    } else if (fs::exists(fs::current_path() / "input templates")) {
        base_dir = fs::current_path();
    } else {
        std::cerr << "Error: Cannot find 'input templates' folder.\n";
        std::cerr << "Please run this program from the project directory.\n";
        return 1;
    }

    fs::path input_dir = base_dir / "input templates";
    fs::path output_dir = base_dir / "output";
    fs::create_directories(output_dir);

    // Set error log path and create messages directory
    g_errors_log_path = base_dir / "messages" / "errors.log";
    fs::create_directories(base_dir / "messages");

    // Set token file path and executable name for credential persistence and error messages
    g_token_file_path = base_dir / "tokens.json";
    g_exe_name = fs::path(argv[0]).filename().string();

    fs::path csv_path = output_dir; // Let download_attendee_report handle CSV selection
    fs::path attendance_template = input_dir / "Attendance Sheet.docx";
    fs::path nametag_template = input_dir / "Name Tag Template.docx";
    // Output paths — will be updated with date/venue prefix after Steps 1-2
    fs::path attendance_output = output_dir / "Attendance Sheet.docx";
    fs::path nametag_output = output_dir / "Name Tag Template.docx";
    fs::path attendance_pdf = output_dir / "Attendance Sheet.pdf";
    fs::path nametag_pdf = output_dir / "Name Tag Template.pdf";

    // Verify templates exist
    if (!fs::exists(attendance_template)) {
        write_error_log(6, "Attendance Sheet template not found.",
            "Path: " + attendance_template.string());
        std::cerr << "Error: Template not found: " << attendance_template.string() << "\n";
        return 1;
    }
    if (!fs::exists(nametag_template)) {
        write_error_log(7, "Name Tag Template not found.",
            "Path: " + nametag_template.string());
        std::cerr << "Error: Template not found: " << nametag_template.string() << "\n";
        return 1;
    }

    // ========================================================================
    // Step 1: Download/obtain attendee report CSV
    // ========================================================================
    std::string event_title;
    std::string event_date;
    if (!download_attendee_report(csv_path, event_title, event_date)) {
        std::cerr << "Error: Could not obtain attendee report CSV.\n";
        return 1;
    }

    // ========================================================================
    // Step 2: Parse CSV and process attendee data
    // ========================================================================
    std::cout << "\n=== Step 2: Process Registration Spreadsheet ===\n\n";

    auto csv_rows = parse_csv(csv_path.string());
    if (csv_rows.size() < 2) {
        write_error_log(8, "CSV file is empty or has no data rows.",
            "Path: " + csv_path.string());
        std::cerr << "Error: CSV file is empty or has no data rows.\n";
        return 1;
    }

    // Skip header row (index 0)
    // Column A (0) = Ticket Type
    // Column F (5) = Attendee's First Name
    // Column G (6) = Attendee's Last Name
    // Column K (10) = Optional Guest's First Name
    // Column L (11) = Optional Guest's Last Name

    // Step 2a: Collect unique ticket types
    std::vector<std::string> ticket_types;
    std::set<std::string> ticket_type_set;

    for (size_t i = 1; i < csv_rows.size(); ++i) {
        std::string tt = csv_rows[i].get(0);
        if (!tt.empty() && ticket_type_set.find(tt) == ticket_type_set.end()) {
            ticket_type_set.insert(tt);
            ticket_types.push_back(tt);
        }
    }

    // Step 2b: Display ticket types and get user choice
    std::sort(ticket_types.begin(), ticket_types.end());
    std::cout << "Available Ticket Types:\n";
    for (size_t i = 0; i < ticket_types.size(); ++i) {
        std::cout << "  " << (i + 1) << ". " << ticket_types[i] << "\n";
    }
    int choice = 0;
    while (true) {
        std::cout << "\nChoose a ticket type (enter number): ";
        std::string choice_str;
        std::getline(std::cin, choice_str);
        { std::string t = trim(choice_str); choice = (!t.empty() && std::all_of(t.begin(), t.end(), ::isdigit)) ? std::stoi(t) : 0; }
        if (choice >= 1 && choice <= (int)ticket_types.size()) break;
        std::cout << "Invalid Response\n";
    }

    std::string ticket_type_selected = ticket_types[choice - 1];
    std::cout << "Selected: " << ticket_type_selected << "\n\n";

    // Step 2c-f: Collect and process attendee names and emails
    std::vector<std::string> attendee_first_names;
    std::vector<std::string> attendee_last_names;
    std::vector<std::string> attendee_emails;

    for (size_t i = 1; i < csv_rows.size(); ++i) {
        if (csv_rows[i].get(0) != ticket_type_selected) continue;

        // Attendee's names (columns F=5, G=6) and email (column E=4)
        std::string first = trim(csv_rows[i].get(5));
        std::string last = trim(csv_rows[i].get(6));
        if (!first.empty()) {
            attendee_first_names.push_back(capitalize(first));
            attendee_last_names.push_back(!last.empty() ? capitalize(last) : "");
            attendee_emails.push_back(trim(csv_rows[i].get(4)));
        }

        // Optional Guest's names (columns K=10, L=11) and email (column J=9)
        std::string guest_first = trim(csv_rows[i].get(10));
        std::string guest_last = trim(csv_rows[i].get(11));
        if (!guest_first.empty()) {
            attendee_first_names.push_back(capitalize(guest_first));
            attendee_last_names.push_back(!guest_last.empty() ? capitalize(guest_last) : "");
            attendee_emails.push_back(trim(csv_rows[i].get(9)));
        }
    }

    // Step 2f: Create attendee list (name + email) and sort
    struct AttendeeEntry { std::string name; std::string email; };
    std::vector<AttendeeEntry> attendee_list;
    for (size_t i = 0; i < attendee_first_names.size(); ++i) {
        attendee_list.push_back({ attendee_first_names[i] + " " + attendee_last_names[i],
                                   attendee_emails[i] });
    }

    // Step 2g: Sort by last name, then first name
    std::sort(attendee_list.begin(), attendee_list.end(),
              [](const AttendeeEntry& a, const AttendeeEntry& b) {
                  auto last_space_a = a.name.rfind(' ');
                  auto last_space_b = b.name.rfind(' ');

                  std::string last_a = (last_space_a != std::string::npos) ? a.name.substr(last_space_a + 1) : a.name;
                  std::string last_b = (last_space_b != std::string::npos) ? b.name.substr(last_space_b + 1) : b.name;

                  std::string first_a = (last_space_a != std::string::npos) ? a.name.substr(0, last_space_a) : "";
                  std::string first_b = (last_space_b != std::string::npos) ? b.name.substr(0, last_space_b) : "";

                  std::string lower_last_a = last_a, lower_last_b = last_b;
                  std::string lower_first_a = first_a, lower_first_b = first_b;
                  std::transform(lower_last_a.begin(), lower_last_a.end(), lower_last_a.begin(),
                                 [](unsigned char c) { return std::tolower(c); });
                  std::transform(lower_last_b.begin(), lower_last_b.end(), lower_last_b.begin(),
                                 [](unsigned char c) { return std::tolower(c); });
                  std::transform(lower_first_a.begin(), lower_first_a.end(), lower_first_a.begin(),
                                 [](unsigned char c) { return std::tolower(c); });
                  std::transform(lower_first_b.begin(), lower_first_b.end(), lower_first_b.begin(),
                                 [](unsigned char c) { return std::tolower(c); });

                  if (lower_last_a != lower_last_b) return lower_last_a < lower_last_b;
                  return lower_first_a < lower_first_b;
              });

    std::vector<std::string> attendee_names;
    std::vector<std::string> attendee_email_list;
    size_t duplicates_removed = 0;
    for (auto& e : attendee_list) {
        // Case-insensitive duplicate check by name
        std::string lower_name = e.name;
        std::transform(lower_name.begin(), lower_name.end(), lower_name.begin(),
                       [](unsigned char c) { return std::tolower(c); });
        bool is_dup = false;
        for (auto& existing : attendee_names) {
            std::string lower_existing = existing;
            std::transform(lower_existing.begin(), lower_existing.end(), lower_existing.begin(),
                           [](unsigned char c) { return std::tolower(c); });
            if (lower_name == lower_existing) { is_dup = true; break; }
        }
        if (is_dup) { ++duplicates_removed; continue; }
        attendee_names.push_back(e.name);
        attendee_email_list.push_back(e.email);
    }
    if (duplicates_removed > 0)
        std::cout << "  Note: " << duplicates_removed << " duplicate name(s) removed.\n";

    std::cout << "Attendees (" << attendee_names.size() << " total):\n";
    for (size_t i = 0; i < attendee_names.size(); ++i) {
        std::cout << "  " << (i + 1) << ". " << attendee_names[i] << "\n";
    }

    // Derive venue name: all but last word of selected ticket type
    // e.g., "Long Island Admission" -> "Long Island"
    std::string venue_name;
    {
        size_t last_space = ticket_type_selected.rfind(' ');
        if (last_space != std::string::npos && last_space > 0)
            venue_name = ticket_type_selected.substr(0, last_space);
        else
            venue_name = ticket_type_selected;
    }

    // Build filename prefix: "YYYY-MM-DD.VenueName." or just "VenueName."
    // event_date holds the CSV filename prefix (e.g. "2026-03-24", "test", or "").
    // venue_name is all but the last word of the selected ticket type.
    {
        std::string file_prefix = (event_date.empty() ? "" : (event_date + ".")) + venue_name + ".";
        attendance_output = output_dir / (file_prefix + "Attendance Sheet.docx");
        nametag_output    = output_dir / (file_prefix + "Name Tag Template.docx");
        attendance_pdf    = output_dir / (file_prefix + "Attendance Sheet.pdf");
        nametag_pdf       = output_dir / (file_prefix + "Name Tag Template.pdf");
    }

    // ========================================================================
    // Step 4: Create Registration List CSV
    // ========================================================================
    std::cout << "\n=== Step 4: Create Registration List ===\n\n";
    {
        std::string file_prefix = (event_date.empty() ? "" : (event_date + ".")) + venue_name + ".";
        fs::path reg_list_path = output_dir / (file_prefix + "Registration List.csv");
        std::ofstream reg_csv(reg_list_path);
        if (!reg_csv.is_open()) {
            write_error_log(9, "Cannot write Registration List CSV.",
                "Path: " + reg_list_path.string());
            std::cerr << "Error: Cannot write Registration List CSV: " << reg_list_path.string() << "\n";
            return 1;
        }
        reg_csv << "Attendee Name,Attendee Email Address\n";
        for (size_t i = 0; i < attendee_names.size(); ++i) {
            reg_csv << csv_escape(attendee_names[i]) << ","
                    << csv_escape(attendee_email_list[i]) << "\n";
        }
        reg_csv.close();
        std::cout << "  Registration List written to: " << reg_list_path.string() << "\n";
    }

    // ========================================================================
    // Step 5: Create Attendance Sheet
    // ========================================================================
    std::cout << "\n=== Step 5: Create Attendance Sheet ===\n\n";

    // Copy template to output
    try {
        fs::copy_file(attendance_template, attendance_output, fs::copy_options::overwrite_existing);
    } catch (const fs::filesystem_error& e) {
        write_error_log(10, "Cannot copy Attendance Sheet template to output.",
            std::string(e.what()));
        std::cerr << "Error: Cannot copy Attendance Sheet template to output.\n";
        std::cerr << "  " << e.what() << "\n";
        std::cerr << "  Please close the file if it is open in another program and try again.\n";
        return 1;
    }
    std::cout << "  Copied template to: " << attendance_output.string() << "\n";

    if (!populate_attendance_sheet(attendance_output, attendee_names, event_title, venue_name)) {
        write_error_log(11, "Failed to populate Attendance Sheet.",
            "Path: " + attendance_output.string());
        std::cerr << "Error: Failed to populate Attendance Sheet.\n";
        return 1;
    }
    std::cout << "  Attendance Sheet populated with " << attendee_names.size() << " names.\n";

    // ========================================================================
    // Step 6: Create Name Tag Template
    // ========================================================================
    std::cout << "\n=== Step 6: Create Name Tag Template ===\n\n";

    // Copy template to output
    try {
        fs::copy_file(nametag_template, nametag_output, fs::copy_options::overwrite_existing);
    } catch (const fs::filesystem_error& e) {
        write_error_log(12, "Cannot copy Name Tag template to output.",
            std::string(e.what()));
        std::cerr << "Error: Cannot copy Name Tag template to output.\n";
        std::cerr << "  " << e.what() << "\n";
        std::cerr << "  Please close the file if it is open in another program and try again.\n";
        return 1;
    }
    std::cout << "  Copied template to: " << nametag_output.string() << "\n";

    if (!populate_name_tags(nametag_output, attendee_names)) {
        write_error_log(13, "Failed to populate Name Tag Template.",
            "Path: " + nametag_output.string());
        std::cerr << "Error: Failed to populate Name Tag Template.\n";
        return 1;
    }

    // ========================================================================
    // Step 7: Convert Attendance Sheet to PDF
    // ========================================================================
    std::cout << "\n=== Step 7: Convert Attendance Sheet to PDF ===\n\n";

    if (!convert_to_pdf(attendance_output, attendance_pdf)) {
        std::cerr << "Warning: PDF conversion failed for Attendance Sheet.\n";
        std::cerr << "  You can open " << attendance_output.string() << " in Word and save as PDF manually.\n";
    }

    // ========================================================================
    // Step 8: Convert Name Tag Template to PDF
    // ========================================================================
    std::cout << "\n=== Step 8: Convert Name Tag Template to PDF ===\n\n";

    if (!convert_to_pdf(nametag_output, nametag_pdf, true)) {
        std::cerr << "Warning: PDF conversion failed for Name Tag Template.\n";
        std::cerr << "  You can open " << nametag_output.string() << " in Word and save as PDF manually.\n";
    }

    // ========================================================================
    // Done
    // ========================================================================
    std::cout << "\n=== Processing Complete ===\n\n";
    std::cout << "Output files:\n";
    std::cout << "  CSV:        " << csv_path.string() << "\n";
    std::cout << "  Attendance: " << attendance_output.string() << "\n";
    std::cout << "  Name Tags:  " << nametag_output.string() << "\n";
    if (fs::exists(attendance_pdf))
        std::cout << "  Attend PDF: " << attendance_pdf.string() << "\n";
    if (fs::exists(nametag_pdf))
        std::cout << "  Tags PDF:   " << nametag_pdf.string() << "\n";

    return 0;
}

int main(int argc, char* argv[]) {
    int result = 0;
    try {
        result = run(argc, argv);
    } catch (const std::exception& e) {
        write_error_log(14, "Unexpected error.", std::string(e.what()));
        std::cerr << "\nUnexpected error: " << e.what() << "\n";
        result = 1;
    } catch (...) {
        write_error_log(15, "Unexpected unknown error.", "No additional details available.");
        std::cerr << "\nUnexpected error occurred.\n";
        result = 1;
    }
    std::cout << "\nPress Enter to exit...";
    std::string dummy;
    std::getline(std::cin, dummy);
    return result;
}
