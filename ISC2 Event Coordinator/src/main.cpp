// ISC2 Event Coordinator - Console Application
// Processes Constant Contact attendee reports and generates attendance sheets and name tags.

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <winhttp.h>
#include <wincrypt.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <comdef.h>
#include <shellapi.h>
#include <gdiplus.h>

#include <pugixml.hpp>
#include <nlohmann/json.hpp>

#include <algorithm>
#include <cctype>
#include <cstdio>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <iomanip>
#include <map>
#include <set>
#include <sstream>
#include <string>
#include <tuple>
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
// Returns empty string if the name doesn't start with a recognizable date.
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

// Pre-fill the console input buffer so the user can edit a suggested value.
// Injects key events into the Windows console input queue before std::getline
// reads. Silently does nothing when stdin is not a real console handle.
static void prefill_console_input(const std::string& text) {
    HANDLE hIn = GetStdHandle(STD_INPUT_HANDLE);
    if (hIn == INVALID_HANDLE_VALUE || hIn == nullptr) return;
    DWORD mode = 0;
    if (!GetConsoleMode(hIn, &mode)) return;  // Not a real console — skip
    std::vector<INPUT_RECORD> records;
    records.reserve(text.size() * 2);
    for (unsigned char ch : text) {
        SHORT vk = VkKeyScanA(static_cast<CHAR>(ch));
        INPUT_RECORD ir = {};
        ir.EventType = KEY_EVENT;
        ir.Event.KeyEvent.wRepeatCount = 1;
        ir.Event.KeyEvent.uChar.AsciiChar = static_cast<CHAR>(ch);
        if (vk != -1) {
            ir.Event.KeyEvent.wVirtualKeyCode = LOBYTE(vk);
            ir.Event.KeyEvent.wVirtualScanCode =
                static_cast<WORD>(MapVirtualKeyA(LOBYTE(vk), MAPVK_VK_TO_VSC));
            BYTE hiVk = HIBYTE(vk);
            if (hiVk & 1) ir.Event.KeyEvent.dwControlKeyState |= SHIFT_PRESSED;
            if (hiVk & 2) ir.Event.KeyEvent.dwControlKeyState |= LEFT_CTRL_PRESSED;
            if (hiVk & 4) ir.Event.KeyEvent.dwControlKeyState |= LEFT_ALT_PRESSED;
        }
        ir.Event.KeyEvent.bKeyDown = TRUE;
        records.push_back(ir);
        ir.Event.KeyEvent.bKeyDown = FALSE;
        records.push_back(ir);
    }
    DWORD written = 0;
    WriteConsoleInputA(hIn, records.data(), static_cast<DWORD>(records.size()), &written);
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

// Error log path (set from run_docreg() after base directory is determined)
static fs::path g_error_log_path;

// Security log path (set from run_docreg() after base directory is determined)
static fs::path g_security_log_path;

static void write_error_log(int error_num, const std::string& description, const std::string& detail) {
    if (g_error_log_path.empty()) return;
    SYSTEMTIME st;
    GetLocalTime(&st);
    char time_buf[32];
    snprintf(time_buf, sizeof(time_buf), "%04d-%02d-%02dT%02d:%02d:%02d",
             st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    std::ofstream log(g_error_log_path, std::ios::app);
    if (!log.is_open()) return;
    char nn[4];
    snprintf(nn, sizeof(nn), "%02d", error_num);
    if (description.empty())
        log << "Error" << nn << ":\n";
    else
        log << "Error" << nn << ": " << description << "\n";
    log << "Time: " << time_buf << "\n";
    log << "Details: " << detail << "\n";
    log << "\n";
}

static void write_security_log(int security_num, const std::string& description, const std::string& detail) {
    if (g_security_log_path.empty()) return;
    SYSTEMTIME st;
    GetLocalTime(&st);
    char time_buf[32];
    snprintf(time_buf, sizeof(time_buf), "%04d-%02d-%02dT%02d:%02d:%02d",
             st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    std::ofstream log(g_security_log_path, std::ios::app);
    if (!log.is_open()) return;
    char nn[4];
    snprintf(nn, sizeof(nn), "%02d", security_num);
    if (description.empty())
        log << "Security" << nn << ":\n";
    else
        log << "Security" << nn << ": " << description << "\n";
    log << "Time: " << time_buf << "\n";
    log << "Details: " << detail << "\n";
    log << "\n";
}

// Forward declarations: shared token I/O defined in the Teams section below,
// used by both Constant Contact and Teams token functions.
static std::string load_tokens_json_plain();
static bool write_tokens_json_plain(const std::string& plain);

struct ConstantContactTokens {
    std::string api_key;
    std::string client_secret;
    std::string access_token;
    std::string refresh_token;
};

static bool save_tokens(const ConstantContactTokens& tok) {
    if (g_token_file_path.empty()) return false;
    json j;
    try { j = json::parse(load_tokens_json_plain()); } catch (...) {}
    j["api_key"]       = tok.api_key;
    j["client_secret"] = tok.client_secret;
    j["access_token"]  = tok.access_token;
    j["refresh_token"] = tok.refresh_token;
    return write_tokens_json_plain(j.dump(2) + "\n");
}

static ConstantContactTokens load_tokens() {
    ConstantContactTokens tok;
    try {
        json j = json::parse(load_tokens_json_plain());
        tok.api_key       = j.value("api_key",       "");
        tok.client_secret = j.value("client_secret", "");
        tok.access_token  = j.value("access_token",  "");
        tok.refresh_token = j.value("refresh_token", "");
    } catch (...) {}
    return tok;
}

static bool refresh_access_token(ConstantContactTokens& tok) {
    if (tok.refresh_token.empty() || tok.api_key.empty() || tok.client_secret.empty()) {
        std::string missing;
        if (tok.refresh_token.empty()) missing += " refresh_token";
        if (tok.client_secret.empty()) missing += " client_secret";
        if (tok.api_key.empty())       missing += " api_key";
        write_security_log(1, "Token refresh failed: missing credentials in tokens.json",
            "Missing fields:" + missing);
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
        write_security_log(2, "Token refresh failed: empty response from authorization server",
            "POST to authz.constantcontact.com/oauth2/default/v1/token returned an empty response.");
        return false;
    }

    try {
        json resp = json::parse(response);
        if (resp.contains("access_token")) {
            tok.access_token  = resp["access_token"].get<std::string>();
            if (resp.contains("refresh_token")) {
                tok.refresh_token = resp["refresh_token"].get<std::string>();
            }
            if (!save_tokens(tok)) return false;
            std::cout << "Token refreshed successfully.\n";
            return true;
        } else {
            write_security_log(3, "Token refresh error: server returned error response",
                resp.value("error", resp.value("error_key", "unknown error")) + ": " +
                resp.value("error_description", resp.value("error_message", "no description")));
            return false;
        }
    } catch (...) {
        write_security_log(4, "Token refresh failed: could not parse authorization server response",
            "Response could not be parsed as JSON.");
        return false;
    }
}

// ============================================================================
// Teams OAuth2 and OpenAI Token Storage
// ============================================================================

struct TeamsTokens {
    std::string tenant_id;
    std::string client_id;
    std::string client_secret;
    std::string access_token;
    std::string refresh_token;
};

// Load the DPAPI-encrypted JSON from tokens.json.
// Returns "{}" if the file does not exist, cannot be read, or cannot be decrypted.
static std::string load_tokens_json_plain() {
    if (g_token_file_path.empty() || !fs::exists(g_token_file_path)) return "{}";
    std::ifstream f(g_token_file_path, std::ios::binary);
    if (!f.is_open()) return "{}";
    std::vector<BYTE> file_data((std::istreambuf_iterator<char>(f)),
                                 std::istreambuf_iterator<char>());
    f.close();
    if (file_data.empty()) return "{}";
    DATA_BLOB in_blob, out_blob;
    in_blob.pbData = file_data.data();
    in_blob.cbData = static_cast<DWORD>(file_data.size());
    if (CryptUnprotectData(&in_blob, nullptr, nullptr, nullptr, nullptr, 0, &out_blob)) {
        std::string plain(reinterpret_cast<const char*>(out_blob.pbData), out_blob.cbData);
        LocalFree(out_blob.pbData);
        return plain;
    }
    // Decryption failed — log and treat as absent. No plaintext fallback.
    write_security_log(7, "tokens.json could not be decrypted: credentials treated as absent.",
        "CryptUnprotectData failed. The file may have been created under a different Windows "
        "user account (DPAPI keys are per-user), or is a legacy plaintext file. "
        "Re-run --setup-cc and/or --setup-teams to re-enter credentials.");
    return "{}"; 
}

// Write JSON to tokens.json, encrypted with DPAPI. Fails hard if encryption is unavailable.
static bool write_tokens_json_plain(const std::string& plain) {
    if (g_token_file_path.empty()) return false;
    std::string mutable_plain = plain;
    DATA_BLOB in_blob, out_blob;
    in_blob.pbData = reinterpret_cast<BYTE*>(mutable_plain.data());
    in_blob.cbData = static_cast<DWORD>(mutable_plain.size());
    if (!CryptProtectData(&in_blob, L"ISC2TokenFile", nullptr, nullptr, nullptr, 0, &out_blob)) {
        DWORD err = GetLastError();
        char err_buf[16];
        snprintf(err_buf, sizeof(err_buf), "0x%08lX", static_cast<unsigned long>(err));
        std::string detail =
            std::string("CryptProtectData failed with Win32 error ") + err_buf + ". "
            "DPAPI (Windows Data Protection API) requires an interactive user session with a "
            "loaded Windows user profile. Common causes: running as SYSTEM or a service account "
            "without a profile, a corrupted or inaccessible user key store "
            "(AppData\\Roaming\\Microsoft\\Protect), or a restrictive group policy. "
            "Credentials were NOT written to disk to prevent plaintext credential exposure.";
        write_error_log(20, "Cannot encrypt tokens.json: DPAPI unavailable.", detail);
        std::cerr << "\nError20: Cannot save credentials — Windows DPAPI encryption failed.\n"
                  << "  Win32 error: " << err_buf << "\n"
                  << "  DPAPI requires an interactive user session with a loaded Windows profile.\n"
                  << "  Common causes: SYSTEM/service account, corrupted key store, group policy.\n"
                  << "  Credentials were NOT written to disk.\n";
        return false;
    }
    std::ofstream f(g_token_file_path, std::ios::binary);
    bool ok = f.is_open();
    if (ok) f.write(reinterpret_cast<const char*>(out_blob.pbData), out_blob.cbData);
    LocalFree(out_blob.pbData);
    return ok;
}

static bool save_teams_tokens(const TeamsTokens& tok) {
    json j;
    try { j = json::parse(load_tokens_json_plain()); } catch (...) {}
    j["teams_tenant_id"]     = tok.tenant_id;
    j["teams_client_id"]     = tok.client_id;
    j["teams_client_secret"] = tok.client_secret;
    j["teams_access_token"]  = tok.access_token;
    j["teams_refresh_token"] = tok.refresh_token;
    return write_tokens_json_plain(j.dump(2) + "\n");
}

static TeamsTokens load_teams_tokens() {
    TeamsTokens tok;
    try {
        json j = json::parse(load_tokens_json_plain());
        tok.tenant_id     = j.value("teams_tenant_id",     "");
        tok.client_id     = j.value("teams_client_id",     "");
        tok.client_secret = j.value("teams_client_secret", "");
        tok.access_token  = j.value("teams_access_token",  "");
        tok.refresh_token = j.value("teams_refresh_token", "");
    } catch (...) {}
    return tok;
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
            write_security_log(5, "HTTP non-200 response from Constant Contact API",
                "HTTP " + std::to_string(status_code) + " from GET " + wide_to_utf8(path));
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

static bool is_yes(const std::string& s) { return s=="y"||s=="Y"||s=="yes"||s=="Yes"; }
static bool is_no (const std::string& s) { return s=="n"||s=="N"||s=="no" ||s=="No";  }

static bool download_attendee_report(fs::path& output_reg_spreadsheet_path,
                                     std::string& out_event_title,
                                     std::string& out_event_date,
                                     bool force_new_download = false) {
    std::cout << "\n=== Choose a Source of Constant Contact Event Registration Data ===\n\n";


    // Determine the output directory: if output_reg_spreadsheet_path is a directory, use it; if it's a file, use its parent
    fs::path output_dir;
    if (fs::is_directory(output_reg_spreadsheet_path)) {
        output_dir = output_reg_spreadsheet_path;
    } else {
        output_dir = output_reg_spreadsheet_path.parent_path();
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

    if (!force_new_download) {
        if (!csv_files.empty()) {
            std::cout << "Choose from existing registration spreadsheets, or create a new registration spreadsheet using the Constant Contact API:\n";
            for (size_t i = 0; i < csv_files.size(); ++i) {
                std::cout << "  " << (i + 1) << ". " << csv_files[i].filename().string() << "\n";
            }
            std::cout << "  " << (csv_files.size() + 1) << ". Create a new registration spreadsheet using the Constant Contact API\n";
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
                output_reg_spreadsheet_path = csv_files[csv_choice - 1];
                std::cout << "Using: " << output_reg_spreadsheet_path.filename().string() << "\n";
                // Extract prefix: everything before "Registration Spreadsheet.csv" in the filename.
                // e.g. "2026-03-24.Registration Spreadsheet.csv" -> "2026-03-24."
                //      "test.Registration Spreadsheet.csv"     -> "test."
                //      "Registration Spreadsheet.csv"          -> ""
                std::string fname = output_reg_spreadsheet_path.filename().string();
                const std::string marker = "Registration Spreadsheet.csv";
                size_t marker_pos = fname.find(marker);
                if (marker_pos != std::string::npos && marker_pos > 0) {
                    // Strip the trailing dot so out_event_date = "2026-03-24" or "test"
                    out_event_date = fname.substr(0, marker_pos - 1); // remove trailing '.'
                }
                // out_event_date remains empty if the file is plain "Registration Spreadsheet.csv"

                // Prompt for event title (no API source available when using existing CSV)
                while (true) {
                    std::cout << "Enter the event title you want to appear on the Attendance Sheet: ";
                    std::getline(std::cin, out_event_title);
                    out_event_title = trim(out_event_title);
                    if (!out_event_title.empty()) break;
                    std::cout << "Invalid Response\n";
                }

                return true;
            }
            // User chose "Create a new registration spreadsheet using the Constant Contact API"
        } else {
            std::cout << "No existing registration spreadsheets were found.\n";
        }
    }
    std::cout << "To create a new registration spreadsheet using the Constant Contact API, you need two credentials:\n";
    std::cout << "  1) API Key (Client ID)\n";
    std::cout << "  2) OAuth2 Access Token\n\n";

    // Load saved tokens
    ConstantContactTokens tokens = load_tokens();
    std::string api_key, token;

    if (!tokens.api_key.empty() && !tokens.access_token.empty()) {
        std::cout << "Saved credentials found (API Key: " << tokens.api_key.substr(0, 8) << "...).\n";
        std::string use_saved;
        while (true) {
            std::cout << "Use saved credentials? (Y/n): ";
            std::getline(std::cin, use_saved);
            use_saved = trim(use_saved);
            if (use_saved.empty() || is_yes(use_saved) || is_no(use_saved)) break;
            std::cout << "Invalid Response\n";
        }
        if (use_saved.empty() || is_yes(use_saved)) {
            api_key = tokens.api_key;
            token   = tokens.access_token;
        }
    }

    if (api_key.empty()) {
        if (!fs::exists(g_token_file_path))
            std::cout << "File " << g_token_file_path.string() << " was not found.\n";
        if (!tokens.api_key.empty())
            std::cout << "Enter your API Key / Client ID (or press Enter to skip API download): ***Warning: Providing this information permanently replaces the current values***\n";
        else
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
    if (!save_tokens(tokens)) return false;

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

    // --- API failure handler: write to messages\error.log and security.log, display structured instructions ---
    auto api_failure = [&](const std::string& error_detail) -> bool {
        write_error_log(1, "Could not obtain the Constant Contact Attendee Report CSV file. Either no registrations were found for the selected event or accessing Constant Contact failed.", error_detail);
        std::string out_folder = output_dir.string();
        std::string log_ref = g_error_log_path.empty()
            ? (out_folder + "\\..\\messages\\error.log")
            : g_error_log_path.string();
        std::cerr << "Error01: Could not obtain the Constant Contact Attendee Report CSV file from Constant Contact. Either no registrations were found for the selected event or accessing Constant Contact failed.\n"
                  << "Contact your Event Coordinator to report this error, including a copy of the file " << log_ref << "\n";
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
        out_event_date = parse_event_date(name);  // YYYY-MM-DD from "Month DD, YYYY ..." name

        // Confirm or override the event title for the Attendance Sheet
        std::cout << "The event title that will appear on the Attendance Sheet is '" << name << "'.\n";
        std::string accept;
        while (true) {
            std::cout << "Accept this event title? (Y/n): ";
            std::getline(std::cin, accept);
            accept = trim(accept);
            if (accept.empty() || is_yes(accept)) {
                out_event_title = name;
                break;
            } else if (is_no(accept)) {
                while (true) {
                    std::cout << "Enter the event title you want to appear on the Attendance Sheet: ";
                    prefill_console_input(name);
                    std::getline(std::cin, out_event_title);
                    out_event_title = trim(out_event_title);
                    if (!out_event_title.empty()) break;
                    std::cout << "Invalid Response\n";
                }
                break;
            }
            std::cout << "Invalid Response\n";
        }
    }

    // Now that we have the event date, set the actual output CSV file path
    {
        std::string prefix = out_event_date.empty() ? "" : (out_event_date + ".");
        output_reg_spreadsheet_path = output_dir / (prefix + "Registration Spreadsheet.csv");
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
    if (fs::exists(output_reg_spreadsheet_path)) {
        std::ofstream test(output_reg_spreadsheet_path, std::ios::app);
        if (!test.is_open()) {
            std::cerr << "Error: CSV file is locked (close Excel first): "
                      << output_reg_spreadsheet_path.string() << "\n";
            std::cout << "Close the file and press Enter to retry...";
            std::string dummy;
            std::getline(std::cin, dummy);
        }
    }
    std::ofstream csv_out(output_reg_spreadsheet_path);
    if (!csv_out.is_open()) {
        write_error_log(5, "Cannot write Registration Spreadsheet CSV.",
            "Path: " + output_reg_spreadsheet_path.string());
        std::cerr << "Error: Cannot create CSV file at " << output_reg_spreadsheet_path.string() << "\n";
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
    std::cout << "Registration data saved to: " << output_reg_spreadsheet_path.string() << "\n";
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
                // If the character immediately before the match is not a space
                // (e.g., the en dash "–" ends in 0x93 with no trailing space in
                // the XML), insert a space so the venue name is not run together
                // with the preceding punctuation.
                std::string replacement_text = venue_name;
                if (pos > 0 && static_cast<unsigned char>(para_text[pos - 1]) != ' ')
                    replacement_text = " " + venue_name;
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
                        runs[i].t_node.text().set((before + replacement_text + after_match).c_str());
                        // Ensure xml:space="preserve" so Word keeps any leading/trailing
                        // whitespace in the replacement text (e.g. the space prepended above).
                        {
                            auto xs = runs[i].t_node.attribute("xml:space");
                            if (!xs) runs[i].t_node.append_attribute("xml:space").set_value("preserve");
                        }
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
            std::cerr << "  Error17: Failed to open document in Word. Syncing the file with OneDrive while there isn't Internet access is a possible cause.\n";
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
// Constant Contact OAuth2 Setup Wizard (--setup-cc)
// ============================================================================

static int run_setup_cc(int argc, char* argv[]) {
    std::cout << "=== ISC2 Event Coordinator - Constant Contact OAuth2 Setup ===\n\n";
    std::cout << "This wizard obtains an access token and refresh token from Constant Contact\n";
    std::cout << "and saves them to tokens.json for use by this program.\n\n";

    // Determine base directory (same logic as run_docreg(), but does not require templates)
    fs::path exe_path = fs::path(argv[0]).parent_path();
    fs::path base_dir;
    if (fs::exists(exe_path / "input templates")) {
        base_dir = exe_path;
    } else if (fs::exists(fs::current_path() / "input templates")) {
        base_dir = fs::current_path();
    } else {
        base_dir = exe_path.empty() ? fs::current_path() : exe_path;
        std::cout << "Note: 'input templates' folder not found. Tokens will be saved to:\n";
        std::cout << "  " << base_dir.string() << "\n\n";
    }

    g_token_file_path   = base_dir / "tokens.json";
    g_exe_name          = fs::path(argv[0]).filename().string();
    g_error_log_path    = base_dir / "messages" / "error.log";
    g_security_log_path = base_dir / "messages" / "security.log";
    fs::create_directories(base_dir / "messages");

    // Load existing tokens as defaults
    ConstantContactTokens tokens = load_tokens();

    // Prompt for API key
    std::string api_key;
    if (!tokens.api_key.empty()) {
        std::cout << "Saved API Key: " << tokens.api_key.substr(0, 8) << "...\n";
        std::cout << "Press Enter to keep it, or type a new API Key / Client ID: ";
        std::getline(std::cin, api_key);
        api_key = trim(api_key);
        if (api_key.empty()) api_key = tokens.api_key;
    } else {
        std::cout << "Enter your API Key / Client ID: ";
        std::getline(std::cin, api_key);
        api_key = trim(api_key);
    }
    if (api_key.empty()) {
        std::cerr << "API Key is required.\n";
        return 1;
    }

    // Prompt for client secret
    std::string client_secret;
    if (!tokens.client_secret.empty()) {
        std::cout << "Saved Client Secret: (masked)\n";
        std::cout << "Press Enter to keep it, or type a new Client Secret: ";
        std::getline(std::cin, client_secret);
        client_secret = trim(client_secret);
        if (client_secret.empty()) client_secret = tokens.client_secret;
    } else {
        std::cout << "Enter your Client Secret: ";
        std::getline(std::cin, client_secret);
        client_secret = trim(client_secret);
    }
    if (client_secret.empty()) {
        std::cerr << "Client Secret is required.\n";
        return 1;
    }

    // Initialize Winsock
    WSADATA wsa_data;
    if (WSAStartup(MAKEWORD(2, 2), &wsa_data) != 0) {
        std::cerr << "Error: Failed to initialize Winsock.\n";
        return 1;
    }

    // Bind to an available local port (8080-8089)
    SOCKET listen_sock = INVALID_SOCKET;
    int port = 0;
    for (int try_port = 8080; try_port <= 8089; ++try_port) {
        SOCKET s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
        if (s == INVALID_SOCKET) continue;
        BOOL reuse = TRUE;
        setsockopt(s, SOL_SOCKET, SO_REUSEADDR,
                   reinterpret_cast<const char*>(&reuse), sizeof(reuse));
        sockaddr_in addr = {};
        addr.sin_family = AF_INET;
        addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
        addr.sin_port = htons(static_cast<u_short>(try_port));
        if (bind(s, reinterpret_cast<SOCKADDR*>(&addr), sizeof(addr)) == 0) {
            listen_sock = s;
            port = try_port;
            break;
        }
        closesocket(s);
    }
    if (listen_sock == INVALID_SOCKET || port == 0) {
        std::cerr << "Error: Could not bind to any local port in range 8080-8089.\n";
        WSACleanup();
        return 1;
    }
    if (listen(listen_sock, 1) != 0) {
        std::cerr << "Error: Failed to listen on port " << port << ".\n";
        closesocket(listen_sock);
        WSACleanup();
        return 1;
    }

    // Generate a random state value for CSRF protection
    std::string state;
    {
        HCRYPTPROV hProv = 0;
        if (CryptAcquireContextW(&hProv, nullptr, nullptr, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)) {
            BYTE state_bytes[16] = {};
            CryptGenRandom(hProv, sizeof(state_bytes), state_bytes);
            CryptReleaseContext(hProv, 0);
            char hex_buf[33] = {};
            for (int i = 0; i < 16; ++i)
                snprintf(hex_buf + i * 2, 3, "%02x", state_bytes[i]);
            state = hex_buf;
        } else {
            state = std::to_string(GetTickCount64());
        }
    }

    std::string redirect_uri = "http://localhost:" + std::to_string(port) + "/callback";

    // Percent-encode the redirect URI for the authorization URL query string
    auto url_encode = [](const std::string& s) -> std::string {
        std::string result;
        for (unsigned char c : s) {
            if (std::isalnum(c) || c == '-' || c == '_' || c == '.' || c == '~')
                result += static_cast<char>(c);
            else {
                char pct[4];
                snprintf(pct, sizeof(pct), "%%%02X", c);
                result += pct;
            }
        }
        return result;
    };

    std::string auth_url =
        "https://authz.constantcontact.com/oauth2/default/v1/authorize"
        "?client_id=" + api_key +
        "&redirect_uri=" + url_encode(redirect_uri) +
        "&response_type=code"
        "&scope=offline_access+campaign_data+contact_data"
        "&state=" + state;

    std::cout << "\nThis setup will use the following OAuth2 redirect URI:\n";
    std::cout << "  " << redirect_uri << "\n\n";
    std::cout << "IMPORTANT: This URI must be registered as a Redirect URI in your\n";
    std::cout << "Constant Contact application settings at:\n";
    std::cout << "  https://app.constantcontact.com/pages/dma/portal/\n\n";
    std::cout << "Press Enter to open your browser for authorization...";
    { std::string dummy; std::getline(std::cin, dummy); }

    ShellExecuteW(nullptr, L"open", utf8_to_wide(auth_url).c_str(),
                  nullptr, nullptr, SW_SHOWNORMAL);

    std::cout << "\nWaiting for authorization callback (timeout: 2 minutes)...\n";
    std::cout << "If your browser did not open, navigate to:\n  " << auth_url << "\n";

    // Wait up to 120 seconds for the browser redirect
    fd_set read_fds;
    FD_ZERO(&read_fds);
    FD_SET(listen_sock, &read_fds);
    TIMEVAL tv = { 120, 0 };
    int sel = select(0, &read_fds, nullptr, nullptr, &tv);
    if (sel <= 0) {
        closesocket(listen_sock);
        WSACleanup();
        std::cerr << (sel == 0 ? "\nError: Timed out waiting for authorization callback.\n"
                               : "\nError: Network error while waiting for callback.\n");
        return 1;
    }

    SOCKET client_sock = accept(listen_sock, nullptr, nullptr);
    closesocket(listen_sock);
    if (client_sock == INVALID_SOCKET) {
        WSACleanup();
        std::cerr << "Error: Failed to accept callback connection.\n";
        return 1;
    }

    // Read the HTTP request (headers only needed)
    std::string request;
    {
        char buf[4096] = {};
        DWORD recv_timeout = 10000;
        setsockopt(client_sock, SOL_SOCKET, SO_RCVTIMEO,
                   reinterpret_cast<const char*>(&recv_timeout), sizeof(recv_timeout));
        int n = recv(client_sock, buf, sizeof(buf) - 1, 0);
        if (n > 0) request.assign(buf, static_cast<size_t>(n));
    }

    // Parse code, state, and error from the request line query string
    std::string code, received_state, error_val;
    {
        auto line_end = request.find("\r\n");
        std::string request_line = (line_end != std::string::npos)
            ? request.substr(0, line_end) : request;
        auto q_pos    = request_line.find('?');
        auto space_pos = request_line.rfind(' ');
        if (q_pos != std::string::npos) {
            size_t q_end = (space_pos != std::string::npos && space_pos > q_pos)
                ? space_pos : request_line.size();
            std::string query = request_line.substr(q_pos + 1, q_end - q_pos - 1);
            std::istringstream qs(query);
            std::string param;
            while (std::getline(qs, param, '&')) {
                auto eq = param.find('=');
                if (eq == std::string::npos) continue;
                std::string key = param.substr(0, eq);
                std::string raw = param.substr(eq + 1);
                // URL-decode value
                std::string val;
                for (size_t i = 0; i < raw.size(); ++i) {
                    if (raw[i] == '%' && i + 2 < raw.size() &&
                        std::isxdigit(static_cast<unsigned char>(raw[i+1])) &&
                        std::isxdigit(static_cast<unsigned char>(raw[i+2]))) {
                        char hex[3] = { raw[i+1], raw[i+2], 0 };
                        val += static_cast<char>(std::strtol(hex, nullptr, 16));
                        i += 2;
                    } else if (raw[i] == '+') {
                        val += ' ';
                    } else {
                        val += raw[i];
                    }
                }
                if (key == "code")  code = val;
                else if (key == "state") received_state = val;
                else if (key == "error") error_val = val;
            }
        }
    }

    // Send response to browser
    bool auth_ok = !code.empty() && received_state == state;
    std::string html_body = auth_ok
        ? "<html><body style=\"font-family:sans-serif;padding:2em\">"
          "<h2 style=\"color:green\">Authorization Successful</h2>"
          "<p>You can close this tab and return to ISC2 Event Coordinator.</p>"
          "</body></html>"
        : "<html><body style=\"font-family:sans-serif;padding:2em\">"
          "<h2 style=\"color:red\">Authorization Failed</h2>"
          "<p>Please return to ISC2 Event Coordinator and try again.</p>"
          "</body></html>";
    std::string http_resp =
        "HTTP/1.1 200 OK\r\n"
        "Content-Type: text/html; charset=utf-8\r\n"
        "Content-Length: " + std::to_string(html_body.size()) + "\r\n"
        "Connection: close\r\n"
        "\r\n" + html_body;
    send(client_sock, http_resp.c_str(), static_cast<int>(http_resp.size()), 0);
    closesocket(client_sock);
    WSACleanup();

    if (!error_val.empty()) {
        std::cerr << "\nError: Authorization denied: " << error_val << "\n";
        return 1;
    }
    if (code.empty()) {
        std::cerr << "\nError: No authorization code received in callback.\n";
        return 1;
    }
    if (received_state != state) {
        std::cerr << "\nError: State mismatch in OAuth callback (possible CSRF).\n";
        return 1;
    }

    std::cout << "\nAuthorization code received. Exchanging for tokens...\n";

    // Build Basic auth header
    std::string credentials = api_key + ":" + client_secret;
    DWORD b64_len = 0;
    CryptBinaryToStringA(reinterpret_cast<const BYTE*>(credentials.c_str()),
                         static_cast<DWORD>(credentials.size()),
                         CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, nullptr, &b64_len);
    std::string b64(b64_len, '\0');
    CryptBinaryToStringA(reinterpret_cast<const BYTE*>(credentials.c_str()),
                         static_cast<DWORD>(credentials.size()),
                         CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, b64.data(), &b64_len);
    b64.resize(b64_len);
    while (!b64.empty() && b64.back() == '\0') b64.pop_back();

    std::string post_body = "grant_type=authorization_code"
                            "&code=" + url_encode(code) +
                            "&redirect_uri=" + url_encode(redirect_uri);

    // POST to token endpoint
    std::string token_response;
    HINTERNET hSession = WinHttpOpen(L"ISC2EventCoordinator/1.0",
                                     WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                     WINHTTP_NO_PROXY_NAME,
                                     WINHTTP_NO_PROXY_BYPASS, 0);
    if (hSession) {
        HINTERNET hConnect = WinHttpConnect(hSession, L"authz.constantcontact.com",
                                            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (hConnect) {
            HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST",
                                                    L"/oauth2/default/v1/token",
                                                    nullptr, WINHTTP_NO_REFERER,
                                                    WINHTTP_DEFAULT_ACCEPT_TYPES,
                                                    WINHTTP_FLAG_SECURE);
            if (hRequest) {
                std::wstring auth_hdr = L"Authorization: Basic " + utf8_to_wide(b64);
                WinHttpAddRequestHeaders(hRequest, auth_hdr.c_str(), (DWORD)-1,
                                         WINHTTP_ADDREQ_FLAG_ADD);
                WinHttpAddRequestHeaders(hRequest,
                                         L"Content-Type: application/x-www-form-urlencoded",
                                         (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);
                if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                                       (LPVOID)post_body.c_str(), (DWORD)post_body.size(),
                                       (DWORD)post_body.size(), 0) &&
                    WinHttpReceiveResponse(hRequest, nullptr)) {
                    DWORD size = 0;
                    do {
                        WinHttpQueryDataAvailable(hRequest, &size);
                        if (size > 0) {
                            std::vector<char> rbuf(size + 1, 0);
                            DWORD downloaded = 0;
                            WinHttpReadData(hRequest, rbuf.data(), size, &downloaded);
                            token_response.append(rbuf.data(), downloaded);
                        }
                    } while (size > 0);
                }
                WinHttpCloseHandle(hRequest);
            }
            WinHttpCloseHandle(hConnect);
        }
        WinHttpCloseHandle(hSession);
    }

    if (token_response.empty()) {
        std::cerr << "Error: Empty response from token endpoint.\n";
        return 1;
    }

    try {
        json resp = json::parse(token_response);
        if (!resp.contains("access_token")) {
            std::cerr << "Error: Token exchange failed.\n  "
                      << resp.value("error", "unknown") << ": "
                      << resp.value("error_description", token_response.substr(0, 300)) << "\n";
            return 1;
        }
        tokens.api_key       = api_key;
        tokens.client_secret = client_secret;
        tokens.access_token  = resp["access_token"].get<std::string>();
        tokens.refresh_token = resp.value("refresh_token", "");
        if (!save_tokens(tokens)) return 1;
        std::cout << "\nSetup complete. Credentials saved to:\n";
        std::cout << "  " << g_token_file_path.string() << "\n\n";
        if (tokens.refresh_token.empty()) {
            std::cout << "Note: No refresh token was returned. Re-run --setup-cc when the access token expires.\n";
        } else {
            std::cout << "A refresh token was saved. The program will renew the access token automatically.\n";
        }
        return 0;
    } catch (const std::exception& e) {
        std::cerr << "Error: Failed to parse token response: " << e.what() << "\n";
        std::cerr << "Response: " << token_response.substr(0, 500) << "\n";
        return 1;
    }
}

// ============================================================================
// Teams Token Refresh
// ============================================================================

static bool refresh_teams_access_token(TeamsTokens& tok) {
    if (tok.refresh_token.empty() || tok.client_id.empty() ||
        tok.client_secret.empty() || tok.tenant_id.empty()) {
        write_security_log(11, "Teams token refresh failed: missing credentials",
            "One or more required fields are missing from Teams token configuration.");
        return false;
    }
    std::cout << "Teams access token expired. Refreshing...\n";

    std::string credentials = tok.client_id + ":" + tok.client_secret;
    DWORD b64_len = 0;
    CryptBinaryToStringA((const BYTE*)credentials.c_str(), (DWORD)credentials.size(),
                         CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, nullptr, &b64_len);
    std::string b64(b64_len, '\0');
    CryptBinaryToStringA((const BYTE*)credentials.c_str(), (DWORD)credentials.size(),
                         CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, b64.data(), &b64_len);
    b64.resize(b64_len);
    while (!b64.empty() && b64.back() == '\0') b64.pop_back();

    std::string body = "grant_type=refresh_token&refresh_token=" + tok.refresh_token
                     + "&scope=OnlineMeetings.Read%20offline_access";
    std::wstring token_host = L"login.microsoftonline.com";
    std::wstring token_path = L"/" + utf8_to_wide(tok.tenant_id) + L"/oauth2/v2.0/token";

    HINTERNET hSession = WinHttpOpen(L"ISC2EventCoordinator/1.0",
                                     WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                     WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return false;
    HINTERNET hConnect = WinHttpConnect(hSession, token_host.c_str(),
                                        INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return false; }
    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", token_path.c_str(),
                                            nullptr, WINHTTP_NO_REFERER,
                                            WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
    if (!hRequest) { WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession); return false; }

    std::wstring auth_w = L"Authorization: Basic " + utf8_to_wide(b64);
    WinHttpAddRequestHeaders(hRequest, auth_w.c_str(), (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);
    WinHttpAddRequestHeaders(hRequest, L"Content-Type: application/x-www-form-urlencoded",
                             (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);
    std::string response;
    if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                           (LPVOID)body.c_str(), (DWORD)body.size(), (DWORD)body.size(), 0) &&
        WinHttpReceiveResponse(hRequest, nullptr)) {
        DWORD size = 0;
        do {
            WinHttpQueryDataAvailable(hRequest, &size);
            if (size > 0) {
                std::vector<char> buf(size + 1, 0); DWORD downloaded = 0;
                WinHttpReadData(hRequest, buf.data(), size, &downloaded);
                response.append(buf.data(), downloaded);
            }
        } while (size > 0);
    }
    WinHttpCloseHandle(hRequest); WinHttpCloseHandle(hConnect); WinHttpCloseHandle(hSession);

    if (response.empty()) {
        write_security_log(12, "Teams token refresh failed: empty response", "");
        return false;
    }
    try {
        json resp = json::parse(response);
        if (resp.contains("access_token")) {
            tok.access_token = resp["access_token"].get<std::string>();
            if (resp.contains("refresh_token"))
                tok.refresh_token = resp["refresh_token"].get<std::string>();
            if (!save_teams_tokens(tok)) return false;
            std::cout << "Teams token refreshed successfully.\n";
            return true;
        }
        write_security_log(13, "Teams token refresh error: server returned error",
            resp.value("error", "unknown") + ": " + resp.value("error_description", ""));
        return false;
    } catch (...) {
        write_security_log(14, "Teams token refresh failed: JSON parse error",
            response.substr(0, 200));
        return false;
    }
}

// ============================================================================
// Microsoft Teams OAuth2 Setup Wizard (--setup-teams)
// ============================================================================

static int run_setup_teams(int argc, char* argv[]) {
    std::cout << "=== ISC2 Event Coordinator - Teams OAuth2 Setup ===\n\n";
    std::cout << "This wizard obtains an access token from Microsoft Azure AD\n";
    std::cout << "and saves it to tokens.json for use by DocAttend.\n\n";
    std::cout << "Prerequisites:\n";
    std::cout << "  - Register an Azure AD app at https://portal.azure.com/\n";
    std::cout << "  - Grant delegated permission: OnlineMeetings.Read\n";
    std::cout << "  - DocAttend only retrieves attendance for meetings you organized.\n\n";

    fs::path exe_path = fs::path(argv[0]).parent_path();
    fs::path base_dir;
    if (fs::exists(exe_path / "input templates"))                 base_dir = exe_path;
    else if (fs::exists(fs::current_path() / "input templates")) base_dir = fs::current_path();
    else                                                           base_dir = exe_path.empty() ? fs::current_path() : exe_path;

    g_token_file_path   = base_dir / "tokens.json";
    g_exe_name          = fs::path(argv[0]).filename().string();
    g_error_log_path    = base_dir / "messages" / "error.log";
    g_security_log_path = base_dir / "messages" / "security.log";
    fs::create_directories(base_dir / "messages");

    TeamsTokens tokens = load_teams_tokens();

    std::string tenant_id;
    if (!tokens.tenant_id.empty()) {
        std::cout << "Saved Tenant ID: " << tokens.tenant_id.substr(0, 8) << "...\n";
        std::cout << "Press Enter to keep it, or type a new Tenant ID: ";
        std::getline(std::cin, tenant_id); tenant_id = trim(tenant_id);
        if (tenant_id.empty()) tenant_id = tokens.tenant_id;
    } else {
        std::cout << "Enter your Azure AD Tenant ID: ";
        std::getline(std::cin, tenant_id); tenant_id = trim(tenant_id);
    }
    if (tenant_id.empty()) { std::cerr << "Tenant ID is required.\n"; return 1; }

    std::string client_id;
    if (!tokens.client_id.empty()) {
        std::cout << "Saved Client ID: " << tokens.client_id.substr(0, 8) << "...\n";
        std::cout << "Press Enter to keep it, or type a new Client ID: ";
        std::getline(std::cin, client_id); client_id = trim(client_id);
        if (client_id.empty()) client_id = tokens.client_id;
    } else {
        std::cout << "Enter your Azure AD Application (Client) ID: ";
        std::getline(std::cin, client_id); client_id = trim(client_id);
    }
    if (client_id.empty()) { std::cerr << "Client ID is required.\n"; return 1; }

    std::string client_secret;
    if (!tokens.client_secret.empty()) {
        std::cout << "Saved Client Secret: (masked)\n";
        std::cout << "Press Enter to keep it, or type a new Client Secret: ";
        std::getline(std::cin, client_secret); client_secret = trim(client_secret);
        if (client_secret.empty()) client_secret = tokens.client_secret;
    } else {
        std::cout << "Enter your Azure AD Client Secret: ";
        std::getline(std::cin, client_secret); client_secret = trim(client_secret);
    }
    if (client_secret.empty()) { std::cerr << "Client Secret is required.\n"; return 1; }

    WSADATA wsa_data;
    if (WSAStartup(MAKEWORD(2, 2), &wsa_data) != 0) {
        std::cerr << "Error: Failed to initialize Winsock.\n"; return 1;
    }

    SOCKET listen_sock = INVALID_SOCKET;
    int port = 0;
    for (int try_port = 8080; try_port <= 8089; ++try_port) {
        SOCKET s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
        if (s == INVALID_SOCKET) continue;
        BOOL reuse = TRUE;
        setsockopt(s, SOL_SOCKET, SO_REUSEADDR, reinterpret_cast<const char*>(&reuse), sizeof(reuse));
        sockaddr_in addr = {};
        addr.sin_family = AF_INET;
        addr.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
        addr.sin_port = htons(static_cast<u_short>(try_port));
        if (bind(s, reinterpret_cast<SOCKADDR*>(&addr), sizeof(addr)) == 0) {
            listen_sock = s; port = try_port; break;
        }
        closesocket(s);
    }
    if (listen_sock == INVALID_SOCKET || port == 0) {
        std::cerr << "Error: Could not bind to any local port in range 8080-8089.\n";
        WSACleanup(); return 1;
    }
    if (listen(listen_sock, 1) != 0) {
        std::cerr << "Error: Failed to listen on port " << port << ".\n";
        closesocket(listen_sock); WSACleanup(); return 1;
    }

    std::string state;
    {
        HCRYPTPROV hProv = 0;
        if (CryptAcquireContextW(&hProv, nullptr, nullptr, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)) {
            BYTE state_bytes[16] = {};
            CryptGenRandom(hProv, sizeof(state_bytes), state_bytes);
            CryptReleaseContext(hProv, 0);
            char hex_buf[33] = {};
            for (int i = 0; i < 16; ++i) snprintf(hex_buf + i * 2, 3, "%02x", state_bytes[i]);
            state = hex_buf;
        } else {
            state = std::to_string(GetTickCount64());
        }
    }

    std::string redirect_uri = "http://localhost:" + std::to_string(port) + "/callback";

    auto url_encode = [](const std::string& s) -> std::string {
        std::string result;
        for (unsigned char c : s) {
            if (std::isalnum(c) || c == '-' || c == '_' || c == '.' || c == '~')
                result += static_cast<char>(c);
            else { char pct[4]; snprintf(pct, sizeof(pct), "%%%02X", c); result += pct; }
        }
        return result;
    };

    std::string auth_url =
        "https://login.microsoftonline.com/" + tenant_id + "/oauth2/v2.0/authorize"
        "?client_id=" + client_id +
        "&redirect_uri=" + url_encode(redirect_uri) +
        "&response_type=code"
        "&scope=OnlineMeetings.Read%20offline_access"
        "&state=" + state;

    std::cout << "\nThis setup uses redirect URI:\n  " << redirect_uri << "\n\n";
    std::cout << "IMPORTANT: Register this URI in your Azure AD app at https://portal.azure.com/\n\n";
    std::cout << "Press Enter to open your browser for authorization...";
    { std::string dummy; std::getline(std::cin, dummy); }

    ShellExecuteW(nullptr, L"open", utf8_to_wide(auth_url).c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    std::cout << "\nWaiting for authorization callback (timeout: 2 minutes)...\n";
    std::cout << "If your browser did not open, navigate to:\n  " << auth_url << "\n";

    fd_set read_fds; FD_ZERO(&read_fds); FD_SET(listen_sock, &read_fds);
    TIMEVAL tv = { 120, 0 };
    int sel = select(0, &read_fds, nullptr, nullptr, &tv);
    if (sel <= 0) {
        closesocket(listen_sock); WSACleanup();
        std::cerr << (sel == 0 ? "\nError: Timed out waiting for authorization callback.\n"
                               : "\nError: Network error while waiting for callback.\n");
        return 1;
    }
    SOCKET client_sock = accept(listen_sock, nullptr, nullptr);
    closesocket(listen_sock);
    if (client_sock == INVALID_SOCKET) {
        WSACleanup(); std::cerr << "Error: Failed to accept callback connection.\n"; return 1;
    }

    std::string request;
    {
        char buf[4096] = {};
        DWORD recv_timeout = 10000;
        setsockopt(client_sock, SOL_SOCKET, SO_RCVTIMEO,
                   reinterpret_cast<const char*>(&recv_timeout), sizeof(recv_timeout));
        int n = recv(client_sock, buf, sizeof(buf) - 1, 0);
        if (n > 0) request.assign(buf, static_cast<size_t>(n));
    }

    std::string code, received_state, error_val;
    {
        auto line_end   = request.find("\r\n");
        std::string request_line = (line_end != std::string::npos)
            ? request.substr(0, line_end) : request;
        auto q_pos     = request_line.find('?');
        auto space_pos = request_line.rfind(' ');
        if (q_pos != std::string::npos) {
            size_t q_end = (space_pos != std::string::npos && space_pos > q_pos)
                ? space_pos : request_line.size();
            std::string query = request_line.substr(q_pos + 1, q_end - q_pos - 1);
            std::istringstream qs(query);
            std::string param;
            while (std::getline(qs, param, '&')) {
                auto eq = param.find('=');
                if (eq == std::string::npos) continue;
                std::string key = param.substr(0, eq);
                std::string raw = param.substr(eq + 1);
                std::string val;
                for (size_t i = 0; i < raw.size(); ++i) {
                    if (raw[i] == '%' && i + 2 < raw.size() &&
                        std::isxdigit(static_cast<unsigned char>(raw[i+1])) &&
                        std::isxdigit(static_cast<unsigned char>(raw[i+2]))) {
                        char hex[3] = { raw[i+1], raw[i+2], 0 };
                        val += static_cast<char>(std::strtol(hex, nullptr, 16));
                        i += 2;
                    } else if (raw[i] == '+') { val += ' ';
                    } else { val += raw[i]; }
                }
                if      (key == "code")  code           = val;
                else if (key == "state") received_state = val;
                else if (key == "error") error_val      = val;
            }
        }
    }

    bool auth_ok = !code.empty() && received_state == state;
    std::string html_body = auth_ok
        ? "<html><body style=\"font-family:sans-serif;padding:2em\">"
          "<h2 style=\"color:green\">Authorization Successful</h2>"
          "<p>You can close this tab and return to ISC2 Event Coordinator.</p>"
          "</body></html>"
        : "<html><body style=\"font-family:sans-serif;padding:2em\">"
          "<h2 style=\"color:red\">Authorization Failed</h2>"
          "<p>Please return to ISC2 Event Coordinator and try again.</p>"
          "</body></html>";
    std::string http_resp =
        "HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\n"
        "Content-Length: " + std::to_string(html_body.size()) + "\r\nConnection: close\r\n\r\n"
        + html_body;
    send(client_sock, http_resp.c_str(), static_cast<int>(http_resp.size()), 0);
    closesocket(client_sock);
    WSACleanup();

    if (!error_val.empty()) {
        std::cerr << "\nError: Authorization denied: " << error_val << "\n"; return 1;
    }
    if (code.empty()) {
        std::cerr << "\nError: No authorization code received in callback.\n"; return 1;
    }
    if (received_state != state) {
        std::cerr << "\nError: State mismatch in OAuth callback (possible CSRF).\n"; return 1;
    }

    std::cout << "\nAuthorization code received. Exchanging for tokens...\n";

    std::string credentials = client_id + ":" + client_secret;
    DWORD b64_len = 0;
    CryptBinaryToStringA(reinterpret_cast<const BYTE*>(credentials.c_str()),
                         static_cast<DWORD>(credentials.size()),
                         CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, nullptr, &b64_len);
    std::string b64(b64_len, '\0');
    CryptBinaryToStringA(reinterpret_cast<const BYTE*>(credentials.c_str()),
                         static_cast<DWORD>(credentials.size()),
                         CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, b64.data(), &b64_len);
    b64.resize(b64_len);
    while (!b64.empty() && b64.back() == '\0') b64.pop_back();

    std::string post_body = "grant_type=authorization_code"
                            "&code=" + url_encode(code) +
                            "&redirect_uri=" + url_encode(redirect_uri) +
                            "&scope=OnlineMeetings.Read%20offline_access";
    std::wstring token_host = L"login.microsoftonline.com";
    std::wstring token_path = L"/" + utf8_to_wide(tenant_id) + L"/oauth2/v2.0/token";
    std::string token_response;

    HINTERNET hSession = WinHttpOpen(L"ISC2EventCoordinator/1.0",
                                     WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                                     WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (hSession) {
        HINTERNET hConnect = WinHttpConnect(hSession, token_host.c_str(),
                                            INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (hConnect) {
            HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"POST", token_path.c_str(),
                                                    nullptr, WINHTTP_NO_REFERER,
                                                    WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
            if (hRequest) {
                std::wstring auth_w2 = L"Authorization: Basic " + utf8_to_wide(b64);
                WinHttpAddRequestHeaders(hRequest, auth_w2.c_str(), (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);
                WinHttpAddRequestHeaders(hRequest,
                    L"Content-Type: application/x-www-form-urlencoded",
                    (DWORD)-1, WINHTTP_ADDREQ_FLAG_ADD);
                if (WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                                       (LPVOID)post_body.c_str(), (DWORD)post_body.size(),
                                       (DWORD)post_body.size(), 0) &&
                    WinHttpReceiveResponse(hRequest, nullptr)) {
                    DWORD size = 0;
                    do {
                        WinHttpQueryDataAvailable(hRequest, &size);
                        if (size > 0) {
                            std::vector<char> buf(size + 1, 0); DWORD downloaded = 0;
                            WinHttpReadData(hRequest, buf.data(), size, &downloaded);
                            token_response.append(buf.data(), downloaded);
                        }
                    } while (size > 0);
                }
                WinHttpCloseHandle(hRequest);
            }
            WinHttpCloseHandle(hConnect);
        }
        WinHttpCloseHandle(hSession);
    }

    if (token_response.empty()) {
        std::cerr << "\nError: Empty response from token endpoint.\n"; return 1;
    }
    json token_json;
    try { token_json = json::parse(token_response); }
    catch (...) { std::cerr << "\nError: Failed to parse token response.\n"; return 1; }

    if (!token_json.contains("access_token")) {
        std::cerr << "\nError: Token exchange failed: "
                  << token_json.value("error", "unknown") << ": "
                  << token_json.value("error_description", token_response.substr(0, 200)) << "\n";
        return 1;
    }

    tokens.tenant_id     = tenant_id;
    tokens.client_id     = client_id;
    tokens.client_secret = client_secret;
    tokens.access_token  = token_json["access_token"].get<std::string>();
    tokens.refresh_token = token_json.value("refresh_token", "");
    if (!save_teams_tokens(tokens)) {
        std::cerr << "\nError: Failed to save Teams tokens.\n"; return 1;
    }
    std::cout << "\nTeams authorization successful!\n";
    std::cout << "Credentials saved to " << g_token_file_path.string() << "\n";
    if (tokens.refresh_token.empty())
        std::cout << "Note: No refresh token received. Re-run --setup-teams periodically.\n";
    return 0;
}

// ============================================================================
// DocAttend: Teams Attendance Download
// ============================================================================

static std::string format_duration_seconds(int total_seconds) {
    int h = total_seconds / 3600;
    int m = (total_seconds % 3600) / 60;
    int s = total_seconds % 60;
    std::ostringstream oss;
    if (h > 0) oss << h << "h ";
    if (h > 0 || m > 0) oss << m << "m ";
    oss << s << "s";
    return oss.str();
}

// Download Teams meeting attendance via Microsoft Graph API.
// Writes {output_dir}/{date}.Teams Attendance.csv (UTF-8, Name,In-Meeting Duration).
// Returns true on success, false if credentials unavailable or API call fails.
static bool download_teams_attendance(const fs::path& output_dir, const std::string& date) {
    // Load credentials
    TeamsTokens tok = load_teams_tokens();
    if (tok.tenant_id.empty() || tok.client_id.empty()) {
        return false;
    }

    // Auto-refresh if no access token
    if (tok.access_token.empty()) {
        try {
            refresh_teams_access_token(tok);
            tok = load_teams_tokens();
        } catch (const std::exception& e) {
            std::cerr << "Failed to get Teams access token: " << e.what() << "\n";
            return false;
        }
    }

    // Helper: URL-encode a UTF-8 string for use in query parameters
    auto url_encode = [](const std::string& s) -> std::wstring {
        std::wstring out;
        for (unsigned char c : s) {
            if (std::isalnum(c) || c == '-' || c == '_' || c == '.' || c == '~') {
                out += static_cast<wchar_t>(c);
            } else {
                wchar_t buf[8];
                swprintf(buf, 8, L"%%%02X", c);
                out += buf;
            }
        }
        return out;
    };

    // Build filter: meetings on the given date
    std::string filter_str = "$filter=startDateTime ge '" + date + "T00:00:00Z'"
                           + " and startDateTime le '" + date + "T23:59:59Z'";
    std::wstring meetings_path = L"/v1.0/me/onlineMeetings?" + url_encode(filter_str);

    // Retry once after token refresh on 401
    auto do_get = [&](const std::wstring& path) -> std::string {
        std::string resp = http_get(L"graph.microsoft.com", path, tok.access_token, "");
        if (resp.find("\"error\"") != std::string::npos &&
            resp.find("401") != std::string::npos) {
            // Refresh and retry
            refresh_teams_access_token(tok);
            tok = load_teams_tokens();
            resp = http_get(L"graph.microsoft.com", path, tok.access_token, "");
        }
        return resp;
    };

    // List meetings
    std::string meetings_json = do_get(meetings_path);
    nlohmann::json meetings_doc;
    try {
        meetings_doc = nlohmann::json::parse(meetings_json);
    } catch (...) {
        std::cerr << "Failed to parse meetings response.\n";
        return false;
    }

    if (meetings_doc.contains("error")) {
        std::string msg = meetings_doc["error"].value("message", meetings_json);
        std::cerr << "Graph API error listing meetings: " << msg << "\n";
        return false;
    }

    auto value_arr = meetings_doc.value("value", nlohmann::json::array());
    if (value_arr.empty()) {
        std::cerr << "No Teams meetings found for " << date << ".\n";
        std::cerr << "Note: Only meetings you organized appear via /me/onlineMeetings.\n";
        return false;
    }

    // Pick meeting (prefill if single)
    std::string meeting_id;
    if (value_arr.size() == 1) {
        meeting_id = value_arr[0].value("id", "");
        std::string subject = value_arr[0].value("subject", "(no subject)");
        std::cout << "Found meeting: " << subject << "\n";
    } else {
        std::cout << "Found " << value_arr.size() << " meetings on " << date << ":\n";
        for (size_t i = 0; i < value_arr.size(); ++i) {
            std::cout << "  " << (i + 1) << ". "
                      << value_arr[i].value("subject", "(no subject)") << "\n";
        }
        std::cout << "Enter meeting number: ";
        std::string choice_str;
        std::getline(std::cin, choice_str);
        int choice = 0;
        try { choice = std::stoi(choice_str); } catch (...) {}
        if (choice < 1 || choice > static_cast<int>(value_arr.size())) {
            std::cerr << "Invalid selection.\n";
            return false;
        }
        meeting_id = value_arr[static_cast<size_t>(choice - 1)].value("id", "");
    }

    if (meeting_id.empty()) {
        std::cerr << "Could not determine meeting ID.\n";
        return false;
    }

    // List attendance reports, pick the most recent (first in list)
    std::wstring reports_path = L"/v1.0/me/onlineMeetings/"
        + utf8_to_wide(meeting_id) + L"/attendanceReports";
    std::string reports_json = do_get(reports_path);
    nlohmann::json reports_doc;
    try {
        reports_doc = nlohmann::json::parse(reports_json);
    } catch (...) {
        std::cerr << "Failed to parse attendance reports response.\n";
        return false;
    }

    if (reports_doc.contains("error")) {
        std::string msg = reports_doc["error"].value("message", reports_json);
        std::cerr << "Graph API error listing attendance reports: " << msg << "\n";
        return false;
    }

    auto reports_arr = reports_doc.value("value", nlohmann::json::array());
    if (reports_arr.empty()) {
        std::cerr << "No attendance reports found for this meeting.\n";
        return false;
    }
    std::string report_id = reports_arr[0].value("id", "");
    if (report_id.empty()) {
        std::cerr << "Could not determine attendance report ID.\n";
        return false;
    }

    // Fetch attendance records
    std::wstring records_path = L"/v1.0/me/onlineMeetings/"
        + utf8_to_wide(meeting_id) + L"/attendanceReports/"
        + utf8_to_wide(report_id) + L"/attendanceRecords";
    std::string records_json = do_get(records_path);
    nlohmann::json records_doc;
    try {
        records_doc = nlohmann::json::parse(records_json);
    } catch (...) {
        std::cerr << "Failed to parse attendance records response.\n";
        return false;
    }

    if (records_doc.contains("error")) {
        std::string msg = records_doc["error"].value("message", records_json);
        std::cerr << "Graph API error fetching attendance records: " << msg << "\n";
        return false;
    }

    auto records_arr = records_doc.value("value", nlohmann::json::array());

    // Write CSV
    fs::create_directories(output_dir);
    fs::path out_path = output_dir / (date + ".Teams Attendance.csv");
    std::ofstream out(out_path);
    if (!out) {
        std::cerr << "Cannot write " << out_path.string() << "\n";
        return false;
    }

    out << "Name,In-Meeting Duration\n";
    for (auto& rec : records_arr) {
        std::string name;
        if (rec.contains("identity") && rec["identity"].contains("displayName")) {
            name = rec["identity"]["displayName"].get<std::string>();
        } else {
            name = rec.value("emailAddress", "Unknown");
        }
        int secs = rec.value("totalAttendanceInSeconds", 0);
        out << csv_escape(name) << "," << csv_escape(format_duration_seconds(secs)) << "\n";
    }
    out.close();

    std::cout << "Teams Attendance saved to: " << out_path.string() << "\n";
    std::cout << "  " << records_arr.size() << " attendee(s) recorded.\n";
    return true;
}

// ============================================================================
// DocAttend: Teams Attendance Parsing
// ============================================================================

struct TeamsAttendee {
    std::string name_raw;       // display name as-is from CSV
    std::string duration_str;   // e.g. "1h 23m 45s"
    int         duration_seconds = 0;
};

// Parse duration strings produced by format_duration_seconds() or Teams-native exports.
// Handles: "Xh Ym Zs", "Xm Ys", "Xs", "H:MM:SS", "M:SS"
static int parse_duration_to_seconds(const std::string& s) {
    if (s.empty()) return 0;

    // "H:MM:SS" or "M:SS" — colon-separated
    if (s.find(':') != std::string::npos) {
        std::istringstream ss(s);
        int a = 0, b = 0, c = 0;
        char sep;
        ss >> a >> sep >> b;
        if (ss >> sep >> c) {
            // H:MM:SS
            return a * 3600 + b * 60 + c;
        } else {
            // M:SS
            return a * 60 + b;
        }
    }

    // "Xh Ym Zs" style — scan tokens
    int total = 0;
    std::istringstream ss(s);
    std::string token;
    while (ss >> token) {
        if (token.empty()) continue;
        char suffix = token.back();
        std::string digits = token.substr(0, token.size() - 1);
        int val = 0;
        try { val = std::stoi(digits); } catch (...) { continue; }
        if (suffix == 'h' || suffix == 'H') total += val * 3600;
        else if (suffix == 'm' || suffix == 'M') total += val * 60;
        else if (suffix == 's' || suffix == 'S') total += val;
    }
    return total;
}

// Parse a Teams Attendance CSV file.
// Supports two formats:
//   1. Teams-native export: UTF-16LE BOM, tab-separated, has "1. Summary" / "2. Participants" sections
//   2. DocAttend-created format: UTF-8, comma-separated, headers Name,In-Meeting Duration
// Returns vector<TeamsAttendee>.
static std::vector<TeamsAttendee> parse_teams_attendance_csv(const fs::path& csv_path) {
    std::vector<TeamsAttendee> result;

    // Read file as binary to detect BOM
    std::ifstream fin(csv_path, std::ios::binary);
    if (!fin) throw std::runtime_error("Cannot open " + csv_path.string());
    std::string raw((std::istreambuf_iterator<char>(fin)),
                     std::istreambuf_iterator<char>());
    fin.close();

    std::string text;
    bool teams_native = false;

    // Detect UTF-16LE BOM (0xFF 0xFE)
    if (raw.size() >= 2 &&
        static_cast<unsigned char>(raw[0]) == 0xFF &&
        static_cast<unsigned char>(raw[1]) == 0xFE) {
        // Convert UTF-16LE → UTF-8
        const wchar_t* wptr = reinterpret_cast<const wchar_t*>(raw.data() + 2);
        size_t wlen = (raw.size() - 2) / 2;
        text = wide_to_utf8(std::wstring(wptr, wlen));
        teams_native = true;
    } else {
        text = raw;
        // Check if it looks like Teams native (starts with "1.")
        if (!text.empty() && text[0] == '1' && text.size() > 1 && text[1] == '.') {
            teams_native = true;
        }
    }

    if (teams_native) {
        // Find "2. Participants" section header line
        size_t part_pos = text.find("2. Participants");
        if (part_pos == std::string::npos) {
            // Nothing to parse
            return result;
        }
        // Advance to next line after "2. Participants"
        size_t line_start = text.find('\n', part_pos);
        if (line_start == std::string::npos) return result;
        ++line_start;

        // Parse tab-separated rows; first row is headers
        // Find column indices for "Full Name" (or "Name") and "In-Meeting Duration"
        int name_col = -1, dur_col = -1;

        auto split_tabs = [](const std::string& line) -> std::vector<std::string> {
            std::vector<std::string> cols;
            std::string cur;
            for (char c : line) {
                if (c == '\t') { cols.push_back(cur); cur.clear(); }
                else if (c != '\r') cur += c;
            }
            cols.push_back(cur);
            return cols;
        };

        bool header_found = false;
        size_t pos = line_start;
        while (pos < text.size()) {
            size_t end = text.find('\n', pos);
            std::string line = (end == std::string::npos)
                ? text.substr(pos)
                : text.substr(pos, end - pos);
            // strip \r
            if (!line.empty() && line.back() == '\r') line.pop_back();
            pos = (end == std::string::npos) ? text.size() : end + 1;

            if (line.empty()) continue;

            auto cols = split_tabs(line);

            if (!header_found) {
                // Find header columns
                for (int i = 0; i < static_cast<int>(cols.size()); ++i) {
                    std::string h = cols[i];
                    // lowercase compare
                    std::string hl; for (char c : h) hl += static_cast<char>(std::tolower(c));
                    if (hl == "full name" || hl == "name") name_col = i;
                    else if (hl.find("in-meeting duration") != std::string::npos ||
                             hl.find("duration") != std::string::npos) dur_col = i;
                }
                if (name_col >= 0 && dur_col >= 0) header_found = true;
                continue;
            }

            // Stop at a blank section divider or next section heading
            if (!cols.empty() && !cols[0].empty() && std::isdigit(cols[0][0]) &&
                cols[0].find('.') != std::string::npos) break;

            if (name_col < 0 || name_col >= static_cast<int>(cols.size())) continue;

            TeamsAttendee att;
            att.name_raw = cols[name_col];
            if (dur_col >= 0 && dur_col < static_cast<int>(cols.size()))
                att.duration_str = cols[dur_col];
            att.duration_seconds = parse_duration_to_seconds(att.duration_str);
            if (!att.name_raw.empty())
                result.push_back(std::move(att));
        }
    } else {
        // DocAttend-created format: UTF-8 CSV, headers Name,In-Meeting Duration
        std::istringstream ss(text);
        std::string line;
        int name_col = -1, dur_col = -1;
        bool header_done = false;

        while (std::getline(ss, line)) {
            if (!line.empty() && line.back() == '\r') line.pop_back();
            if (line.empty()) continue;

            // Simple CSV split (no quoted fields with commas expected in names/durations)
            std::vector<std::string> cols;
            std::string cur;
            bool in_quote = false;
            for (char c : line) {
                if (c == '"') { in_quote = !in_quote; }
                else if (c == ',' && !in_quote) { cols.push_back(cur); cur.clear(); }
                else { cur += c; }
            }
            cols.push_back(cur);

            if (!header_done) {
                for (int i = 0; i < static_cast<int>(cols.size()); ++i) {
                    std::string h = cols[i];
                    std::string hl; for (char c : h) hl += static_cast<char>(std::tolower(c));
                    if (hl == "name") name_col = i;
                    else if (hl.find("duration") != std::string::npos) dur_col = i;
                }
                header_done = true;
                continue;
            }

            if (name_col < 0 || name_col >= static_cast<int>(cols.size())) continue;

            TeamsAttendee att;
            att.name_raw = cols[name_col];
            if (dur_col >= 0 && dur_col < static_cast<int>(cols.size()))
                att.duration_str = cols[dur_col];
            att.duration_seconds = parse_duration_to_seconds(att.duration_str);
            if (!att.name_raw.empty())
                result.push_back(std::move(att));
        }
    }

    return result;
}

// ============================================================================
// DocAttend: Registration Matching
// ============================================================================

struct RegistrationEntry {
    std::string first;
    std::string last;
    std::string email;
    std::string isc2_member_id;
};

// Build a lookup map from (lowercase_first, lowercase_last) -> RegistrationEntry.
// Processes both primary attendee (cols 5,6,4,8) and optional guest (cols 10,11,9,13).
// Skips row 0 (header row).
static std::map<std::pair<std::string,std::string>, RegistrationEntry>
build_registration_lookup(const std::vector<CsvRow>& rows) {
    std::map<std::pair<std::string,std::string>, RegistrationEntry> lookup;

    auto to_lower = [](std::string s) -> std::string {
        for (char& c : s) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
        return s;
    };

    // Extract name variants from a field that may contain a parenthesized nickname,
    // e.g. "Adedoyin (Doyin)" -> {"Adedoyin (Doyin)", "Adedoyin", "Doyin"}.
    // Returns the original field plus any base/nickname alternatives.
    auto name_variants = [](const std::string& name) -> std::vector<std::string> {
        std::vector<std::string> v;
        v.push_back(name);
        size_t pa = name.find('(');
        if (pa != std::string::npos) {
            size_t pb = name.find(')', pa);
            std::string base = name.substr(0, pa);
            // Trim trailing whitespace from base
            while (!base.empty() && (base.back() == ' ' || base.back() == '\t'))
                base.pop_back();
            if (!base.empty()) v.push_back(base);
            if (pb != std::string::npos) {
                std::string nick = name.substr(pa + 1, pb - pa - 1);
                // Trim whitespace from nick
                size_t s = nick.find_first_not_of(" \t");
                size_t e2 = nick.find_last_not_of(" \t");
                if (s != std::string::npos) nick = nick.substr(s, e2 - s + 1);
                if (!nick.empty()) v.push_back(nick);
            }
        }
        return v;
    };

    auto add_entry = [&](const std::string& first, const std::string& last,
                         const std::string& email, const std::string& isc2_id) {
        if (first.empty() && last.empty()) return;
        RegistrationEntry e;
        e.first        = first;
        e.last         = last;
        e.email        = email;
        e.isc2_member_id = isc2_id;
        // Store all combinations of first/last name variants so nicknames match
        for (const auto& f : name_variants(first))
            for (const auto& l : name_variants(last))
                lookup.emplace(std::make_pair(to_lower(f), to_lower(l)), e);
    };

    for (size_t i = 1; i < rows.size(); ++i) {
        const CsvRow& row = rows[i];
        // Primary attendee: first=col5, last=col6, email=col4, isc2=col8
        add_entry(row.get(5), row.get(6), row.get(4), row.get(8));
        // Optional guest: first=col10, last=col11, email=col9, isc2=col13
        add_entry(row.get(10), row.get(11), row.get(9), row.get(13));
    }

    return lookup;
}

// Tokenize a raw Teams display name into candidate (first, last) pairs.
// Handles:
//   - Standard "First Last" and "First M. Last" (middle initial skip)
//   - "Last, First" comma format — tries both orderings
//   - Parenthesized real names: "Role_Name (First Last)" or "Name (External)"
//   - Common label words filtered: External, Unverified, Verified, Guest, Dr, Mr, etc.
static std::vector<std::pair<std::string,std::string>>
extract_name_candidates(const std::string& raw_name) {
    std::vector<std::pair<std::string,std::string>> candidates;

    auto add = [&](const std::string& first, const std::string& last) {
        if (!first.empty() && !last.empty())
            candidates.push_back({first, last});
    };

    // Labels to suppress during tokenization (case-insensitive exact match)
    static const std::set<std::string> labels = {
        "external", "unverified", "verified", "guest",
        "dr", "mr", "mrs", "ms", "prof"
    };

    // Split a string into name words (alpha/hyphen/apostrophe only),
    // filtering out known label words.
    auto tokenize = [&](const std::string& s) -> std::vector<std::string> {
        std::vector<std::string> words;
        std::string cur;
        auto flush = [&]() {
            if (cur.empty()) return;
            std::string lc;
            for (char ch : cur)
                lc += static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
            if (!labels.count(lc)) words.push_back(cur);
            cur.clear();
        };
        for (unsigned char c : s) {
            if (std::isalpha(c) || c == '-' || c == '\'') cur += static_cast<char>(c);
            else flush();
        }
        flush();
        return words;
    };

    // Build (first,last) pairs from a word list.
    // Also skips single-letter middle initials to produce (first, last) directly.
    // For exactly three tokens, also try a compound last name: First + "Second Third".
    auto build_pairs = [&](const std::vector<std::string>& words) {
        for (size_t i = 0; i + 1 < words.size(); ++i) {
            add(words[i], words[i + 1]);
            // Middle initial skip: if next token is a single letter, also try word after it
            if (words[i + 1].size() == 1 && i + 2 < words.size())
                add(words[i], words[i + 2]);
        }
        if (words.size() == 3) {
            add(words[0], words[1] + " " + words[2]);
        }
    };

    // 1. Comma format: "Last, First" — only when the part before the comma is a single
    // name word (e.g. "Sanford, Anthony").  If part1 has two or more words the comma
    // separates credentials/suffix (e.g. "Greg Caporale, CISSP (External)") and we
    // fall through so that cases 2/3 can extract the full name correctly.
    size_t comma_pos = raw_name.find(',');
    if (comma_pos != std::string::npos) {
        auto w1 = tokenize(raw_name.substr(0, comma_pos));
        if (w1.size() == 1) {
            auto w2 = tokenize(raw_name.substr(comma_pos + 1));
            if (!w2.empty()) {
                add(w2[0], w1[0]); // "Last, First" interpretation (try first)
                add(w1[0], w2[0]); // "First, Last" interpretation (try second)
            }
            return candidates;
        }
        // w1 has 2+ words: fall through to paren/standard processing
    }

    // 2. Parenthesized content: extract and try first, then try the outside
    size_t pa = raw_name.find('(');
    if (pa != std::string::npos) {
        size_t pb = raw_name.find(')', pa);
        if (pb != std::string::npos) {
            std::string inside  = raw_name.substr(pa + 1, pb - pa - 1);
            std::string outside = raw_name.substr(0, pa);
            if (pb + 1 < raw_name.size()) outside += raw_name.substr(pb + 1);
            build_pairs(tokenize(inside));   // e.g. "Phillip Martin" from "Role (Phillip Martin)"
            build_pairs(tokenize(outside));  // e.g. "Ray Sitorus" from "Ray Sitorus (External)"
            return candidates;
        }
    }

    // 3. Standard tokenization with middle-initial skip
    build_pairs(tokenize(raw_name));
    return candidates;
}

// Try to match a raw Teams display name to a registration entry.
// Returns pointer into lookup on match, nullptr if no match found.
static const RegistrationEntry* match_name_to_registration(
    const std::string& raw_name,
    const std::map<std::pair<std::string,std::string>, RegistrationEntry>& lookup)
{
    auto to_lower = [](std::string s) -> std::string {
        for (char& c : s) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
        return s;
    };

    auto candidates = extract_name_candidates(raw_name);
    for (auto& [first, last] : candidates) {
        auto key = std::make_pair(to_lower(first), to_lower(last));
        auto it = lookup.find(key);
        if (it != lookup.end()) return &it->second;
    }
    return nullptr;
}



// ============================================================================
// DocAttend: Teams Diagnostic CSV
// ============================================================================

struct TeamsDiagRow {
    std::string name_raw;
    int         duration_seconds = 0;
    std::string disposition;
    std::string candidate_first;  // semicolon-separated list of all tried first names
    std::string candidate_last;   // semicolon-separated list of all tried last names
    std::string registered_name;  // "First Last" from registration spreadsheet, or "None found"
    std::string matched_email;
    std::string matched_id;
};

// Re-reads the original Teams Attendance CSV and writes a copy with seven
// diagnostic columns (H-N) appended to the right of every data row.
// Handles both UTF-16LE Teams-native (tab-separated, sectioned) and
// UTF-8 DocAttend-created (comma-separated) formats.
// Output is always UTF-8 with BOM, comma-separated.
static void write_teams_diagnostic_csv(
    const fs::path& source_csv,
    const fs::path& diag_path,
    const std::vector<TeamsDiagRow>& diag_rows)
{
    // Name -> diagnostic data lookup
    std::map<std::string, const TeamsDiagRow*> by_name;
    for (auto& d : diag_rows) by_name[d.name_raw] = &d;

    // Read source file as binary
    std::ifstream fin(source_csv, std::ios::binary);
    if (!fin) return;
    std::string raw((std::istreambuf_iterator<char>(fin)), {});
    fin.close();

    std::string text;
    bool teams_native = false;
    char sep = ',';

    if (raw.size() >= 2 &&
        static_cast<unsigned char>(raw[0]) == 0xFF &&
        static_cast<unsigned char>(raw[1]) == 0xFE) {
        const wchar_t* wptr = reinterpret_cast<const wchar_t*>(raw.data() + 2);
        size_t wlen = (raw.size() - 2) / 2;
        text = wide_to_utf8(std::wstring(wptr, wlen));
        teams_native = true;
        sep = '\t';
    } else {
        text = raw;
        if (!text.empty() && text[0] == '1' && text.size() > 1 && text[1] == '.') {
            teams_native = true;
            sep = '\t';
        }
    }

    std::ofstream out(diag_path);
    if (!out) return;
    out << "\xEF\xBB\xBF"; // UTF-8 BOM for Excel

    // Split a row into fields using sep (handles basic quoting for comma-sep)
    auto split_fields = [sep](const std::string& line) -> std::vector<std::string> {
        std::vector<std::string> cols;
        std::string cur;
        bool in_q = false;
        for (char c : line) {
            if (c == '"') in_q = !in_q;
            else if (c == sep && !in_q) { cols.push_back(cur); cur.clear(); }
            else cur += c;
        }
        cols.push_back(cur);
        return cols;
    };

    // Write original fields (comma-escaped) followed by 6 diagnostic fields
    auto write_row = [&](const std::vector<std::string>& orig, const TeamsDiagRow* d) {
        for (size_t i = 0; i < orig.size(); ++i) {
            if (i > 0) out << ',';
            out << csv_escape(orig[i]);
        }
        out << ',';
        if (d) {
            out << d->duration_seconds             << ','
                << csv_escape(d->disposition)      << ','
                << csv_escape(d->candidate_first)  << ','
                << csv_escape(d->candidate_last)   << ','
                << csv_escape(d->registered_name)  << ','
                << csv_escape(d->matched_email)    << ','
                << csv_escape(d->matched_id);
        } else {
            out << ",,,,,,"; // 7 empty diagnostic cells
        }
        out << '\n';
    };

    std::istringstream ss(text);
    std::string line;
    bool in_participants = !teams_native; // non-native: all lines are participant rows
    bool header_done     = false;
    int  name_col        = -1;

    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();

        // Before the "2. Participants" section in Teams-native format
        if (teams_native && !in_participants) {
            if (line.find("2. Participants") != std::string::npos)
                in_participants = true;
            write_row(split_fields(line), nullptr);
            continue;
        }

        // Blank lines: preserve as empty rows
        if (line.empty()) {
            out << '\n';
            continue;
        }

        auto cols = split_fields(line);

        // Column header row
        if (!header_done) {
            for (int i = 0; i < static_cast<int>(cols.size()); ++i) {
                std::string hl;
                for (char c : cols[i])
                    hl += static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
                if (hl == "full name" || hl == "name") { name_col = i; break; }
            }
            for (size_t i = 0; i < cols.size(); ++i) {
                if (i > 0) out << ',';
                out << csv_escape(cols[i]);
            }
            out << ",Duration (seconds),Disposition,Candidate First Name,"
                   "Candidate Last Name,Matched Registered Name,Matched Email,Matched ISC2 Member ID\n";
            header_done = true;
            continue;
        }

        // Data row: look up by name
        std::string name;
        if (name_col >= 0 && name_col < static_cast<int>(cols.size()))
            name = cols[name_col];
        auto it = name.empty() ? by_name.end() : by_name.find(name);
        write_row(cols, it != by_name.end() ? it->second : nullptr);
    }
}

// ============================================================================
// DocAttend: Main Pipeline
// ============================================================================

static int run_docattend(int argc, char* argv[]) {
    std::cout << "=== ISC2 Event Coordinator: DocAttend ===\n";
    std::cout << "Generates a CPE Eligible Attendees report.\n\n";

    // Determine base directory
    fs::path exe_path = fs::path(argv[0]).parent_path();
    fs::path base_dir;
    if      (fs::exists(exe_path / "Registration Spreadsheets"))           base_dir = exe_path;
    else if (fs::exists(fs::current_path() / "Registration Spreadsheets")) base_dir = fs::current_path();
    else if (fs::exists(exe_path / "input templates"))                     base_dir = exe_path;
    else if (fs::exists(fs::current_path() / "input templates"))           base_dir = fs::current_path();
    else                                                                    base_dir = exe_path.empty() ? fs::current_path() : exe_path;

    fs::path reg_dir       = base_dir / "Registration Spreadsheets";
    fs::path teams_dir     = base_dir / "Teams Attendance";
    fs::path output_dir    = base_dir / "output";
    fs::create_directories(reg_dir);
    g_token_file_path   = base_dir / "tokens.json";
    g_exe_name          = fs::path(argv[0]).filename().string();
    g_error_log_path    = base_dir / "messages" / "error.log";
    g_security_log_path = base_dir / "messages" / "security.log";
    fs::create_directories(base_dir / "messages");

    // -------------------------------------------------------------------------
    // Scan Registration Spreadsheets, pick event date
    // -------------------------------------------------------------------------
    std::vector<fs::path> reg_files;
    for (auto& entry : fs::directory_iterator(reg_dir)) {
        if (entry.is_regular_file()) {
            std::string fname = entry.path().filename().string();
            if (fname.find("Registration Spreadsheet") != std::string::npos &&
                entry.path().extension() == ".csv") {
                reg_files.push_back(entry.path());
            }
        }
    }
    std::sort(reg_files.begin(), reg_files.end());

    bool just_downloaded = false;

    if (reg_files.empty()) {
        std::cout << "No Registration Spreadsheet CSV files found in:\n  "
                  << reg_dir.string() << "\n";
        std::cout << "Download a Registration Spreadsheet using the Constant Contact API? [Y/n]: ";
        std::string dl_yn;
        std::getline(std::cin, dl_yn);
        dl_yn = trim(dl_yn);
        if (dl_yn.empty() || is_yes(dl_yn)) {
            fs::path dl_path = reg_dir;
            std::string dummy_title, dummy_date;
            if (!download_attendee_report(dl_path, dummy_title, dummy_date, true)) {
                return 1;
            }
            just_downloaded = true;
            for (auto& entry : fs::directory_iterator(reg_dir)) {
                if (entry.is_regular_file()) {
                    std::string fname = entry.path().filename().string();
                    if (fname.find("Registration Spreadsheet") != std::string::npos &&
                        entry.path().extension() == ".csv") {
                        reg_files.push_back(entry.path());
                    }
                }
            }
            std::sort(reg_files.begin(), reg_files.end());
        }
        if (reg_files.empty()) {
            std::cerr << "No Registration Spreadsheet CSV files found.\n";
            return 1;
        }
    }

    std::string chosen_date;
    fs::path chosen_reg_file;

    if (reg_files.size() == 1) {
        chosen_reg_file = reg_files[0];
        // Extract date prefix (YYYY-MM-DD) from filename
        std::string stem = reg_files[0].stem().string();
        if (stem.size() >= 10) chosen_date = stem.substr(0, 10);
        std::cout << "Using registration file: " << chosen_reg_file.filename().string() << "\n";
    } else {
        std::cout << "Available event dates:\n";
        for (size_t i = 0; i < reg_files.size(); ++i) {
            std::string stem = reg_files[i].stem().string();
            std::string date_part = stem.size() >= 10 ? stem.substr(0, 10) : stem;
            std::cout << "  " << (i + 1) << ". " << date_part << "\n";
        }
        std::cout << "  " << (reg_files.size() + 1)
                  << ". Create a new registration spreadsheet using the Constant Contact API\n";
        int choice = 0;
        while (choice < 1 || choice > static_cast<int>(reg_files.size() + 1)) {
            std::cout << "Enter event number: ";
            std::string choice_str;
            std::getline(std::cin, choice_str);
            choice_str = trim(choice_str);
            try { choice = std::stoi(choice_str); } catch (...) { choice = 0; }
            if (choice < 1 || choice > static_cast<int>(reg_files.size() + 1))
                std::cout << "  Please enter a number between 1 and " << (reg_files.size() + 1) << ".\n";
        }
        if (choice == static_cast<int>(reg_files.size() + 1)) {
            fs::path dl_path = reg_dir;
            std::string dummy_title, dummy_date;
            if (!download_attendee_report(dl_path, dummy_title, dummy_date, true)) {
                return 1;
            }
            chosen_reg_file = dl_path;
            just_downloaded = true;
        } else {
            chosen_reg_file = reg_files[static_cast<size_t>(choice - 1)];
        }
        std::string stem = chosen_reg_file.stem().string();
        if (stem.size() >= 10) chosen_date = stem.substr(0, 10);
    }

    if (chosen_date.empty()) {
        std::cerr << "Could not determine event date from filename.\n";
        return 1;
    }

    // Offer to refresh Registration Spreadsheet from Constant Contact
    if (!just_downloaded) {
        std::cout << "Refresh Registration Spreadsheet using the Constant Contact API? [y/N]: ";
        std::string ref_yn;
        std::getline(std::cin, ref_yn);
        ref_yn = trim(ref_yn);
        if (is_yes(ref_yn)) {
            std::string dummy_title, dummy_date;
            download_attendee_report(chosen_reg_file, dummy_title, dummy_date, true);
        }
    }

    // Parse registration spreadsheet
    std::vector<CsvRow> reg_rows = parse_csv(chosen_reg_file.string());
    auto lookup = build_registration_lookup(reg_rows);
    std::cout << "Loaded " << (reg_rows.size() > 1 ? reg_rows.size() - 1 : 0) << " registered attendee(s).\n\n";

    // Output accumulators
    // cpe rows: first, last, email, cpe, status, isc2_member_id
    std::vector<std::vector<std::string>> cpe_list;
    // Track CPE entries by (lower_first, lower_last) to deduplicate
    std::set<std::pair<std::string,std::string>> cpe_seen;

    auto to_lower_str = [](std::string s) -> std::string {
        for (char& c : s) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
        return s;
    };

    auto add_cpe = [&](const std::string& first, const std::string& last,
                       const std::string& email, const std::string& status,
                       const std::string& isc2_id) {
        auto key = std::make_pair(to_lower_str(first), to_lower_str(last));
        if (cpe_seen.insert(key).second) {
            cpe_list.push_back({capitalize(first), capitalize(last), email, "Yes", status, isc2_id});
        }
    };

    // -------------------------------------------------------------------------
    // Check / download Teams Attendance CSV
    // -------------------------------------------------------------------------
    fs::create_directories(teams_dir);

    // Scan for existing Teams Attendance CSV files matching the selected event date
    std::vector<fs::path> teams_files;
    for (auto& entry : fs::directory_iterator(teams_dir)) {
        if (entry.is_regular_file()) {
            std::string fname = entry.path().filename().string();
            if (fname.find("Teams Attendance") != std::string::npos &&
                entry.path().extension() == ".csv" &&
                fname.size() >= chosen_date.size() &&
                fname.substr(0, chosen_date.size()) == chosen_date) {
                teams_files.push_back(entry.path());
            }
        }
    }
    std::sort(teams_files.begin(), teams_files.end());

    fs::path teams_csv;
    bool have_teams_csv = false;

    if (teams_files.size() == 1) {
        teams_csv = teams_files[0];
        have_teams_csv = true;
        std::cout << "Using Teams Attendance file: " << teams_csv.filename().string() << "\n";
    } else if (teams_files.size() > 1) {
        std::cout << "Choose from existing Teams Attendance files, or create a new Teams Attendance file using the Microsoft Graph API:\n";
        for (size_t i = 0; i < teams_files.size(); ++i)
            std::cout << "  " << (i + 1) << ". " << teams_files[i].filename().string() << "\n";
        std::cout << "  " << (teams_files.size() + 1)
                  << ". Create a new Teams Attendance file using the Microsoft Graph API\n\n";
        int choice = 0;
        while (choice < 1 || choice > static_cast<int>(teams_files.size() + 1)) {
            std::cout << "Choose an option (enter number): ";
            std::string choice_str;
            std::getline(std::cin, choice_str);
            choice_str = trim(choice_str);
            try { choice = std::stoi(choice_str); } catch (...) { choice = 0; }
            if (choice < 1 || choice > static_cast<int>(teams_files.size() + 1))
                std::cout << "  Please enter a number between 1 and " << (teams_files.size() + 1) << ".\n";
        }
        if (choice == static_cast<int>(teams_files.size() + 1)) {
            std::cout << "Attempting to download Teams Attendance via Microsoft Graph API...\n";
            bool downloaded = download_teams_attendance(teams_dir, chosen_date);
            if (downloaded) {
                teams_csv = teams_dir / (chosen_date + ".Teams Attendance.csv");
                have_teams_csv = fs::exists(teams_csv);
            } else {
                std::cout << "\nYour Microsoft Teams credentials are not configured.\n";
                std::cout << "Run EventCoordinator --setup-teams once to configure your credentials, then rerun EventCoordinator --DocAttend to continue processing by automatically downloading a Teams Attendance file using those credentials.\n\n";
                std::cout << "Alternatively, use the following procedure to manually download a Teams Attendance file, then rerun EventCoordinator --DocAttend to continue processing with that file.\n\n";
                std::cout << "1. Open Microsoft Teams for Windows.\n";
                std::cout << "2. Select Chat from the left side of the Teams app.\n";
                std::cout << "3. Select the Meeting chats filter.\n";
                std::cout << "4. Select the meeting of interest from the list presented (You must be an Organizer or Co-organizer of the meeting).\n";
                std::cout << "5. Select Attendance from the top of the Teams app.\n";
                std::cout << "6. Select Download from the right side of the Teams app.\n";
                std::cout << "7. Save the downloaded file to the EventCoordinator folder 'Teams Attendance', using the filename format YYYY-MM-DD.Teams Attendance.csv, for example " << chosen_date << ".Teams Attendance.csv.\n\n";
                return 0;
            }
        } else {
            teams_csv = teams_files[static_cast<size_t>(choice - 1)];
            have_teams_csv = true;
        }
    } else {
        // No files found — attempt API download directly
        std::cout << "Attempting to download Teams Attendance via Microsoft Graph API...\n";
        bool downloaded = download_teams_attendance(teams_dir, chosen_date);
        teams_csv = teams_dir / (chosen_date + ".Teams Attendance.csv");
        if (downloaded) {
            have_teams_csv = fs::exists(teams_csv);
        } else {
            std::cout << "\nYour Microsoft Teams credentials are not configured.\n";
            std::cout << "Run EventCoordinator --setup-teams once to configure your credentials, then rerun EventCoordinator --DocAttend to continue processing by automatically downloading a Teams Attendance file using those credentials.\n\n";
            std::cout << "Alternatively, use the following procedure to manually download a Teams Attendance file, then rerun EventCoordinator --DocAttend to continue processing with that file.\n\n";
            std::cout << "1. Open Microsoft Teams for Windows.\n";
            std::cout << "2. Select Chat from the left side of the Teams app.\n";
            std::cout << "3. Select the meeting of interest from the list presented (You must be an Organizer or Co-organizer of the meeting).\n";
            std::cout << "4. Select Attendance from the top of the Teams app.\n";
            std::cout << "5. Select Download from the right side of the Teams app.\n";
            std::cout << "6. Save the downloaded file to the EventCoordinator folder 'Teams Attendance', using the filename format YYYY-MM-DD.Teams Attendance.csv, for example " << chosen_date << ".Teams Attendance.csv.\n\n";
            return 0;
        }
    }

    // -------------------------------------------------------------------------
    // Process Teams Attendance (>= 90 minutes = 5400 seconds)
    // -------------------------------------------------------------------------
    if (have_teams_csv && fs::exists(teams_csv)) {
        std::vector<TeamsAttendee> attendees;
        try {
            attendees = parse_teams_attendance_csv(teams_csv);
        } catch (const std::exception& e) {
            std::cerr << "Warning: Could not parse Teams Attendance CSV: " << e.what() << "\n";
        }

        int count_cpe = 0, count_unreg = 0, count_short = 0, count_no_id = 0;
        std::vector<TeamsDiagRow> diag_rows;

        std::cout << "\nTeams Attendance: " << attendees.size() << " attendee(s) in CSV\n";

        for (auto& att : attendees) {
            TeamsDiagRow dr;
            dr.name_raw         = att.name_raw;
            dr.duration_seconds = att.duration_seconds;

            // Build semicolon-separated lists of all candidate (first, last) pairs
            auto cands = extract_name_candidates(att.name_raw);
            for (size_t ci = 0; ci < cands.size(); ++ci) {
                if (ci > 0) { dr.candidate_first += "; "; dr.candidate_last += "; "; }
                dr.candidate_first += cands[ci].first;
                dr.candidate_last  += cands[ci].second;
            }

            if (att.duration_seconds < 5400) {
                dr.disposition      = "Skipped: under 90 minutes";
                dr.registered_name  = "N/A";
                ++count_short;
            } else {
                const RegistrationEntry* reg = match_name_to_registration(att.name_raw, lookup);
                if (reg) {
                    dr.registered_name = reg->first + " " + reg->last;
                    dr.matched_email   = reg->email;
                    dr.matched_id      = reg->isc2_member_id;
                    if (!reg->isc2_member_id.empty()) {
                        add_cpe(reg->first, reg->last, reg->email, "Registered", reg->isc2_member_id);
                        dr.disposition = "CPE eligible";
                        ++count_cpe;
                    } else {
                        dr.disposition = "Registered, no ISC2 Member ID on file";
                        ++count_no_id;
                    }
                } else {
                    dr.disposition     = "Unregistered: no match in Registration Spreadsheet";
                    dr.registered_name = "None found";
                    ++count_unreg;
                }
            }
            diag_rows.push_back(std::move(dr));
        }

        std::cout << "  " << count_cpe << " CPE-eligible, "
                  << count_unreg << " unregistered, "
                  << count_short << " skipped (under 90 min)";
        if (count_no_id > 0)
            std::cout << ", " << count_no_id << " registered/no ISC2 Member ID";
        std::cout << "\n";

        // Write diagnostic CSV (original columns + 7 diagnostic columns H-N)
        fs::path msg_dir  = base_dir / "messages";
        fs::create_directories(msg_dir);
        fs::path diag_path = msg_dir / (chosen_date + ".Teams Attendance Diagnostics.csv");
        write_teams_diagnostic_csv(teams_csv, diag_path, diag_rows);
        std::cout << "  Diagnostics: " << diag_path.string() << "\n";
    }

    // -------------------------------------------------------------------------
    // Attendance Sheet — Interactive Console Workflow
    // -------------------------------------------------------------------------

    // Pre-printed signed names (registered, non-virtual attendees)
    std::cout << "\n--- Pre-printed Attendance Sheet Names of Registered Attendees with a Known ISC2 Member ID ---\n";
    std::cout << "If there are pre-printed names on an Attendance Sheet that are not prompted for, it's because these attendees have either not registered (they were manually added to the Attendance Sheet), have not provided an ISC2 Member ID, or both.\n";
    std::cout << "For each registered in-person attendee, indicate whether they signed\n";
    std::cout << "the " << chosen_date << " Attendance Sheet.\n\n";
    {
        struct InPersonEntry {
            std::string ticket_type;
            std::string first;
            std::string last;
            std::string email;
            std::string isc2_id;
        };
        std::vector<InPersonEntry> in_person;
        std::set<std::pair<std::string,std::string>> seen;

        auto collect = [&](const std::string& first, const std::string& last,
                           const std::string& email, const std::string& isc2_id,
                           const std::string& ticket_type) {
            if (first.empty() || last.empty() || email.empty() || isc2_id.empty()) return;
            auto key = std::make_pair(to_lower_str(first), to_lower_str(last));
            if (!seen.insert(key).second) return;
            in_person.push_back({ticket_type, first, last, email, isc2_id});
        };

        for (size_t i = 1; i < reg_rows.size(); ++i) {
            const CsvRow& row = reg_rows[i];
            if (row.get(0) == "Virtual Admission") continue;
            collect(row.get(5), row.get(6), row.get(4), row.get(8), row.get(0));
            collect(row.get(10), row.get(11), row.get(9), row.get(13), row.get(0));
        }

        // Sort by ticket type, then last name, then first name
        std::sort(in_person.begin(), in_person.end(),
            [&](const InPersonEntry& a, const InPersonEntry& b) {
                std::string at = to_lower_str(a.ticket_type);
                std::string bt = to_lower_str(b.ticket_type);
                if (at != bt) return at < bt;
                std::string al = to_lower_str(a.last);
                std::string bl = to_lower_str(b.last);
                if (al != bl) return al < bl;
                return to_lower_str(a.first) < to_lower_str(b.first);
            });

        for (auto& e : in_person) {
            std::cout << "Has ";
            if (!e.ticket_type.empty()) std::cout << capitalize(e.ticket_type) << "/ ";
            std::cout << capitalize(e.first) << " " << capitalize(e.last)
                      << " signed the " << chosen_date << " Attendance Sheet? [y/N]: ";
            std::string yn;
            std::getline(std::cin, yn);
            yn = trim(yn);
            if (is_yes(yn))
                add_cpe(e.first, e.last, e.email, "Registered", e.isc2_id);
        }
    }

    // Page 2 walk-ins (unregistered attendees)
    std::cout << "\n--- Walk-in (Unregistered) Attendees (Page 2) ---\n";
    std::cout << "Enter the attendee information printed on page 2 of each venue's Attendance Sheet.\n";
    while (true) {
        std::cout << "Name on page 2 (blank = done): ";
        std::string full_name;
        std::getline(std::cin, full_name);
        full_name = trim(full_name);
        if (full_name.empty()) break;

        // Split on last space to get first / last
        std::string walkin_first, walkin_last;
        size_t sp = full_name.rfind(' ');
        if (sp != std::string::npos) {
            walkin_first = trim(full_name.substr(0, sp));
            walkin_last  = trim(full_name.substr(sp + 1));
        } else {
            walkin_first = full_name;
        }

        std::cout << "  Email: ";
        std::string walkin_email;
        std::getline(std::cin, walkin_email);
        walkin_email = trim(walkin_email);
        std::cout << "  ISC2 Member ID: ";
        std::string walkin_isc2;
        std::getline(std::cin, walkin_isc2);
        walkin_isc2 = trim(walkin_isc2);

        if (walkin_first.empty() || walkin_email.empty() || walkin_isc2.empty()) {
            std::cout << "  Name, email, and ISC2 Member ID are all required -- skipping.\n";
            continue;
        }

        add_cpe(walkin_first, walkin_last, walkin_email, "Unregistered", walkin_isc2);
        std::cout << "  -> Added.\n";
    }

    // -------------------------------------------------------------------------
    // Write output CSV files
    // -------------------------------------------------------------------------
    fs::create_directories(output_dir);

    // Sort CPE list by last name, then first name
    std::sort(cpe_list.begin(), cpe_list.end(),
        [](const std::vector<std::string>& a, const std::vector<std::string>& b) {
            if (a[1] != b[1]) return a[1] < b[1];
            return a[0] < b[0];
        });

    // CPE Eligible Attendees
    fs::path cpe_out_path = output_dir / (chosen_date + ".CPE Eligible Attendees.csv");
    {
        std::ofstream f(cpe_out_path);
        if (!f) { std::cerr << "Cannot write " << cpe_out_path.string() << "\n"; return 1; }
        f << "Registrant First Name,Registrant Last Name,Registrant Email Address,"
             "CPE,Registration Status,ISC2 Member ID\n";
        for (auto& row : cpe_list) {
            f << csv_escape(row[0]) << "," << csv_escape(row[1]) << ","
              << csv_escape(row[2]) << "," << csv_escape(row[3]) << ","
              << csv_escape(row[4]) << "," << row[5] << "\n";
        }
    }

    std::cout << "\nDone!\n";
    std::cout << "  CPE Eligible:    " << cpe_list.size()   << " attendee(s) -> " << cpe_out_path.string()   << "\n";
    return 0;
}

// ============================================================================
// Main Program
// ============================================================================

static int run_docreg(int argc, char* argv[]) {
    std::cout << "=== ISC2 Event Coordinator ===\n";
    std::cout << "Generates an Attendance Sheet (sign-in sheet), Attendee Name Tag Labels, and a Registration List (attendee names and email addresses) from Constant Contact event registration data.\n\n";

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
    fs::path reg_spreadsheets_dir = base_dir / "Registration Spreadsheets";
    fs::create_directories(reg_spreadsheets_dir);

    // Set error log path and create messages directory
    g_error_log_path    = base_dir / "messages" / "error.log";
    g_security_log_path = base_dir / "messages" / "security.log";
    fs::create_directories(base_dir / "messages");

    // Set token file path and executable name for credential persistence and error messages
    g_token_file_path = base_dir / "tokens.json";
    g_exe_name = fs::path(argv[0]).filename().string();

    fs::path reg_spreadsheet_path = reg_spreadsheets_dir; // Let download_attendee_report handle CSV selection
    fs::path attendance_template = input_dir / "Attendance Sheet.docx";
    fs::path nametag_template = input_dir / "Name Tag Template.docx";
    // Output paths — will be updated with date/venue prefix later
    fs::path attendance_output; // = output_dir / "Attendance Sheet.docx";
    fs::path nametag_output;    // = output_dir / "Name Tag Template.docx";
    fs::path attendance_pdf;    // = output_dir / "Attendance Sheet.pdf";
    fs::path nametag_pdf;       // = output_dir / "Name Tag Template.pdf";
    fs::path reg_list_path;     // = output_dir / "Registration List.csv"

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
    // Download/obtain attendee report CSV
    // ========================================================================
    std::string event_title;
    std::string event_date;
    if (!download_attendee_report(reg_spreadsheet_path, event_title, event_date)) {
        return 1;
    }

    // ========================================================================
    // Parse CSV and process attendee data
    // ========================================================================
    std::cout << "\n=== Process the Registration Spreadsheet ===\n\n";

    auto csv_rows = parse_csv(reg_spreadsheet_path.string());
    if (csv_rows.size() < 2) {
        write_error_log(8, "CSV file is empty or has no data rows.",
            "Path: " + reg_spreadsheet_path.string());
        std::cerr << "Error: CSV file is empty or has no data rows.\n";
        return 1;
    }

    // Skip header row (index 0)
    // Column A (0) = Ticket Type
    // Column F (5) = Attendee's First Name
    // Column G (6) = Attendee's Last Name
    // Column K (10) = Optional Guest's First Name
    // Column L (11) = Optional Guest's Last Name

    // Collect unique ticket types
    std::vector<std::string> ticket_types;
    std::set<std::string> ticket_type_set;

    for (size_t i = 1; i < csv_rows.size(); ++i) {
        std::string tt = csv_rows[i].get(0);
        if (!tt.empty() && ticket_type_set.find(tt) == ticket_type_set.end()) {
            ticket_type_set.insert(tt);
            ticket_types.push_back(tt);
        }
    }

    // Display ticket types and get user choice
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

    // Collect and process attendee names and emails
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

    // Create attendee list (name + email) and sort
    struct AttendeeEntry { std::string name; std::string email; };
    std::vector<AttendeeEntry> attendee_list;
    for (size_t i = 0; i < attendee_first_names.size(); ++i) {
        attendee_list.push_back({ attendee_first_names[i] + " " + attendee_last_names[i],
                                   attendee_emails[i] });
    }

    // Sort by last name, then first name
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
    std::cout << "Attendees: (" << attendee_names.size() << " total";
    if (duplicates_removed > 0)
        std::cout << ". " << duplicates_removed << " duplicate name(s) removed";
    std::cout << ")\n";
    for (size_t i = 0; i < attendee_names.size(); ++i) {
        std::cout << "  " << (i + 1) << ". " << attendee_names[i] << "\n";
    }

    // Derive venue name: all but last word of selected ticket type
    // e.g., "Long Island Admission" -> "Long Island"
    // ticket_type_prefix is used in output filenames and is always the ticket-derived value.
    // venue_name may be overridden by the user and is used only in the document placeholder.
    std::string venue_name;
    {
        size_t last_space = ticket_type_selected.rfind(' ');
        if (last_space != std::string::npos && last_space > 0)
            venue_name = ticket_type_selected.substr(0, last_space);
        else
            venue_name = ticket_type_selected;
    }
    const std::string ticket_type_prefix = venue_name;  // locked before any user override

    // Confirm or override the event venue for the Attendance Sheet
    {
        std::cout << "\nThe event venue that will appear on the Attendance Sheet is '" << venue_name << "', and will be written as '" << venue_name << " Venue'.\n";
        std::string accept;
        while (true) {
            std::cout << "Accept this event venue? (Y/n): ";
            std::getline(std::cin, accept);
            accept = trim(accept);
            if (accept.empty() || is_yes(accept)) {
                break;
            } else if (is_no(accept)) {
                std::string input;
                while (true) {
                    std::cout << "Enter the event venue you want to appear on the Attendance Sheet: ";
                    prefill_console_input(venue_name);
                    std::getline(std::cin, input);
                    input = trim(input);
                    if (!input.empty()) break;
                    std::cout << "Invalid Response\n";
                }
                venue_name = input;
                std::cout << "The event venue that will appear on the Attendance Sheet is '" << venue_name << "', and will be written as '" << venue_name << " Venue'.\n";
                break;
            }
            std::cout << "Invalid Response\n";
        }
    }

    // Build filename prefix: "YYYY-MM-DD.TicketTypeSelected." or just "TicketTypeSelected."
    // event_date holds the CSV filename prefix (e.g. "2026-03-24", "test", or "").
    // ticket_type_prefix is all but the last word of the selected ticket type (never overridden).
    {
        std::string file_prefix = (event_date.empty() ? "" : (event_date + ".")) + ticket_type_prefix + ".";
        attendance_output = output_dir / (file_prefix + "Attendance Sheet.docx");
        nametag_output    = output_dir / (file_prefix + "Name Tag Template.docx");
        attendance_pdf    = output_dir / (file_prefix + "Attendance Sheet.pdf");
        nametag_pdf       = output_dir / (file_prefix + "Name Tag Template.pdf");
        reg_list_path     = output_dir / (file_prefix + "Registration List.csv");
    }

    // ========================================================================
    // Create Registration List CSV
    // ========================================================================
    std::cout << "\n=== Create Registration List ===\n\n";
    {
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
    // Create Attendance Sheet
    // ========================================================================
    std::cout << "\n=== Create Attendance Sheet ===\n\n";

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
    // Create Name Tag Template
    // ========================================================================
    std::cout << "\n=== Create Name Tag Template ===\n\n";

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
    // Convert Attendance Sheet to PDF
    // ========================================================================
    std::cout << "\n=== Convert Attendance Sheet to PDF ===\n\n";

    const bool attendance_pdf_created = convert_to_pdf(attendance_output, attendance_pdf);
    if (!attendance_pdf_created) {
        std::cerr << "  Error18: PDF conversion failed for Attendance Sheet.\n";
        std::cerr << "  You can open " << attendance_output.string() << " in Word and save as PDF manually.\n";
    }

    // ========================================================================
    // Convert Name Tag Template to PDF
    // ========================================================================
    std::cout << "\n=== Convert Name Tag Template to PDF ===\n\n";

    const bool nametag_pdf_created = convert_to_pdf(nametag_output, nametag_pdf, true);
    if (!nametag_pdf_created) {
        std::cerr << "  Error19: PDF conversion failed for Name Tag Template.\n";
        std::cerr << "  You can open " << nametag_output.string() << " in Word and save as PDF manually.\n";
    }

    // ========================================================================
    // Done
    // ========================================================================
    std::cout << "\n=== Processing Complete ===\n\n";
    std::cout << "Output files:\n";
    std::cout << "  Registration Spreadsheet CSV: " << reg_spreadsheet_path.string() << "\n";
    std::cout << "  Registration List CSV:        " << reg_list_path.string() << "\n";
    std::cout << "  Attendance Sheet DOCX:        " << attendance_output.string() << "\n";
    std::cout << "  Name Tag Template DOCX:       " << nametag_output.string() << "\n";
    if (attendance_pdf_created)
        std::cout << "  Attendance Sheet PDF:         " << attendance_pdf.string() << "\n";
    if (nametag_pdf_created)
        std::cout << "  Name Tag Template PDF:        " << nametag_pdf.string() << "\n";

    return 0;
}

// Forward declarations for functions defined after this point.
static int run_setup_teams(int argc, char* argv[]);
static int run_docattend(int argc, char* argv[]);

int main(int argc, char* argv[]) {
    int result = 0;

    // Parse command-line flag
    std::string mode_flag;
    for (int i = 1; i < argc; ++i) {
        std::string arg = argv[i];
        if (arg == "--DocReg"      || arg == "--DocAttend"  ||
            arg == "--setup-cc"    || arg == "--setup-teams") {
            mode_flag = arg;
            break;
        }
    }

    // Interactive menu when no flag supplied
    if (mode_flag.empty()) {
        std::cout << "=== ISC2 Event Coordinator ===\n\n";
        std::cout << "  1. DocReg    - Attendance Sheet, Name Tags, Registration List\n";
        std::cout << "  2. DocAttend - CPE Eligible Attendees report\n\n";
        int choice = 0;
        while (true) {
            std::cout << "Choose an option (enter number): ";
            std::string resp;
            std::getline(std::cin, resp);
            std::string t = trim(resp);
            choice = (!t.empty() && std::all_of(t.begin(), t.end(), ::isdigit))
                     ? std::stoi(t) : 0;
            if (choice == 1 || choice == 2) break;
            std::cout << "Invalid Response\n";
        }
        mode_flag = (choice == 1) ? "--DocReg" : "--DocAttend";
    }

    try {
        if      (mode_flag == "--DocReg")        result = run_docreg(argc, argv);
        else if (mode_flag == "--DocAttend")      result = run_docattend(argc, argv);
        else if (mode_flag == "--setup-cc")       result = run_setup_cc(argc, argv);
        else if (mode_flag == "--setup-teams")    result = run_setup_teams(argc, argv);
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