#define _CRT_SECURE_NO_WARNINGS
#define _WINSOCK_DEPRECATED_NO_WARNINGS
#define WIN32_LEAN_AND_MEAN
#define NTDDI_VERSION NTDDI_WIN10
#define _WIN32_WINNT _WIN32_WINNT_WIN10

#pragma warning(disable: 4996)

#include <iostream>
#include <string>
#include <vector>
#include <map>
#include <sstream>
#include <fstream>
#include <iomanip>
#include <winsock2.h>
#include <ws2tcpip.h>
#include <sapi.h>
#include <sphelper.h>
#include <thread>
#include <mutex>
#include <memory>

#pragma comment(lib, "ws2_32.lib")
#pragma comment(lib, "sapi.lib")

#define MAX_BUFFER 8192
#define PORT 8000

typedef struct lame_global_flags* (*lame_init_t)(void);
typedef int (*lame_set_num_channels_t)(struct lame_global_flags*, int);
typedef int (*lame_set_in_samplerate_t)(struct lame_global_flags*, int);
typedef int (*lame_set_mode_t)(struct lame_global_flags*, int);
typedef int (*lame_init_params_t)(struct lame_global_flags*);
typedef int (*lame_encode_buffer_t)(struct lame_global_flags*, const short int*, const short int*, int, unsigned char*, int);
typedef int (*lame_encode_flush_t)(struct lame_global_flags*, unsigned char*, int);
typedef void (*lame_close_t)(struct lame_global_flags*);

static lame_init_t lame_init = NULL;
static lame_set_num_channels_t lame_set_num_channels = NULL;
static lame_set_in_samplerate_t lame_set_in_samplerate = NULL;
static lame_set_mode_t lame_set_mode = NULL;
static lame_init_params_t lame_init_params = NULL;
static lame_encode_buffer_t lame_encode_buffer = NULL;
static lame_encode_flush_t lame_encode_flush = NULL;
static lame_close_t lame_close = NULL;
static HMODULE hLameDll = NULL;

class ComInitializer {
public:
    ComInitializer() : m_hr(S_OK) {
        m_hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
        if (FAILED(m_hr)) {
            std::cerr << "[ERROR] CoInitializeEx failed: 0x" << std::hex << m_hr << std::endl;
        }
    }
    
    ~ComInitializer() {
        if (SUCCEEDED(m_hr)) {
            CoUninitialize();
        }
    }
    
    bool IsInitialized() const { return SUCCEEDED(m_hr); }
    HRESULT GetHResult() const { return m_hr; }

private:
    HRESULT m_hr;
};

bool LoadLameLibrary() {
    hLameDll = LoadLibraryA("lame_enc.dll");
    if (!hLameDll) {
        hLameDll = LoadLibraryA("libmp3lame.dll");
    }
    if (!hLameDll) {
        std::cerr << "[INFO] Cannot load lame_enc.dll or libmp3lame.dll" << std::endl;
        return false;
    }
    
    lame_init = (lame_init_t)GetProcAddress(hLameDll, "lame_init");
    lame_set_num_channels = (lame_set_num_channels_t)GetProcAddress(hLameDll, "lame_set_num_channels");
    lame_set_in_samplerate = (lame_set_in_samplerate_t)GetProcAddress(hLameDll, "lame_set_in_samplerate");
    lame_set_mode = (lame_set_mode_t)GetProcAddress(hLameDll, "lame_set_mode");
    lame_init_params = (lame_init_params_t)GetProcAddress(hLameDll, "lame_init_params");
    lame_encode_buffer = (lame_encode_buffer_t)GetProcAddress(hLameDll, "lame_encode_buffer");
    lame_encode_flush = (lame_encode_flush_t)GetProcAddress(hLameDll, "lame_encode_flush");
    lame_close = (lame_close_t)GetProcAddress(hLameDll, "lame_close");
    
    if (!lame_init || !lame_set_num_channels || !lame_set_in_samplerate ||
        !lame_init_params || !lame_encode_buffer || !lame_encode_flush || !lame_close) {
        std::cerr << "[ERROR] Cannot get lame function addresses" << std::endl;
        FreeLibrary(hLameDll);
        hLameDll = NULL;
        return false;
    }
    
    std::cout << "[INFO] LAME library loaded successfully" << std::endl;
    return true;
}

void UnloadLameLibrary() {
    if (hLameDll) {
        FreeLibrary(hLameDll);
        hLameDll = NULL;
        std::cout << "[INFO] LAME library unloaded" << std::endl;
    }
}

std::wstring Utf8ToWide(const std::string& utf8) {
    if (utf8.empty()) return L"";
    int size = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, NULL, 0);
    std::wstring result(size - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, &result[0], size);
    return result;
}

std::string WideToUtf8(const std::wstring& wide) {
    if (wide.empty()) return "";
    int size = WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, NULL, 0, NULL, NULL);
    std::string result(size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, &result[0], size, NULL, NULL);
    return result;
}

std::string UrlDecode(const std::string& encoded) {
    std::string decoded;
    char ch;
    int i, ii;
    for (i = 0; i < (int)encoded.length(); i++) {
        if (encoded[i] == '%') {
            if (i + 2 >= (int)encoded.length()) break;
            std::string hex = encoded.substr(i + 1, 2);
            ii = (int)strtol(hex.c_str(), NULL, 16);
            ch = static_cast<char>(ii);
            decoded += ch;
            i += 2;
        } else if (encoded[i] == '+') {
            decoded += ' ';
        } else {
            decoded += encoded[i];
        }
    }
    return decoded;
}

std::string JsonEscape(const std::string& s) {
    std::ostringstream oss;
    for (size_t i = 0; i < s.length(); i++) {
        char c = s[i];
        switch (c) {
            case '\\': oss << "\\\\"; break;
            case '"': oss << "\\\""; break;
            case '\n': oss << "\\n"; break;
            case '\r': oss << "\\r"; break;
            case '\t': oss << "\\t"; break;
            default:
                if ((unsigned char)c < 32) {
                    oss << "\\u00" << std::hex << std::setw(2) << std::setfill('0') << (int)(unsigned char)c;
                } else {
                    oss << c;
                }
        }
    }
    return oss.str();
}

std::vector<std::pair<std::wstring, std::wstring>> GetVoiceListInternal() {
    std::vector<std::pair<std::wstring, std::wstring>> voices;
    HRESULT hr;
    
    std::cout << "[DEBUG] GetVoiceListInternal: Starting..." << std::endl;
    
    ISpVoice* pVoice = NULL;
    hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] CoCreateInstance(CLSID_SpVoice) failed: 0x" << std::hex << hr << std::endl;
        return voices;
    }
    std::cout << "[DEBUG] ISpVoice created successfully" << std::endl;
    
    ISpObjectTokenCategory* pCategory = NULL;
    hr = CoCreateInstance(CLSID_SpObjectTokenCategory, NULL, CLSCTX_ALL, IID_ISpObjectTokenCategory, (void**)&pCategory);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] CoCreateInstance(CLSID_SpObjectTokenCategory) failed: 0x" << std::hex << hr << std::endl;
        pVoice->Release();
        return voices;
    }
    std::cout << "[DEBUG] ISpObjectTokenCategory created successfully" << std::endl;
    
    hr = pCategory->SetId(SPCAT_VOICES, FALSE);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] ISpObjectTokenCategory::SetId failed: 0x" << std::hex << hr << std::endl;
        pCategory->Release();
        pVoice->Release();
        return voices;
    }
    std::cout << "[DEBUG] Category ID set to SPCAT_VOICES" << std::endl;
    
    IEnumSpObjectTokens* pEnum = NULL;
    hr = pCategory->EnumTokens(NULL, NULL, &pEnum);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] ISpObjectTokenCategory::EnumTokens failed: 0x" << std::hex << hr << std::endl;
        pCategory->Release();
        pVoice->Release();
        return voices;
    }
    std::cout << "[DEBUG] EnumTokens succeeded" << std::endl;
    
    ISpObjectToken* pToken = NULL;
    ULONG ulFetched = 0;
    int voiceCount = 0;
    
    while (SUCCEEDED(pEnum->Next(1, &pToken, &ulFetched)) && ulFetched == 1) {
        voiceCount++;
        std::cout << "[DEBUG] Found voice #" << voiceCount << std::endl;
        
        WCHAR* pId = NULL;
        if (SUCCEEDED(pToken->GetId(&pId))) {
            std::wstring id = pId;
            CoTaskMemFree(pId);
            std::cout << "[DEBUG] Voice ID: " << WideToUtf8(id) << std::endl;
            
            WCHAR* pDesc = NULL;
            if (SUCCEEDED(SpGetDescription(pToken, &pDesc))) {
                std::wstring desc = pDesc;
                CoTaskMemFree(pDesc);
                std::cout << "[DEBUG] Voice Description: " << WideToUtf8(desc) << std::endl;
                voices.push_back(std::make_pair(id, desc));
            } else {
                std::cerr << "[WARNING] SpGetDescription failed for voice" << std::endl;
            }
        } else {
            std::cerr << "[WARNING] ISpObjectToken::GetId failed" << std::endl;
        }
        pToken->Release();
    }
    
    std::cout << "[DEBUG] Total voices found: " << voices.size() << std::endl;
    
    pEnum->Release();
    pCategory->Release();
    pVoice->Release();
    
    return voices;
}

std::vector<BYTE> TextToWavInternal(const std::wstring& text, const std::wstring& voiceId) {
    std::vector<BYTE> wavData;
    HRESULT hr;
    
    std::cout << "[DEBUG] TextToWavInternal: Starting..." << std::endl;
    std::cout << "[DEBUG] Text length: " << text.length() << std::endl;
    if (!voiceId.empty()) {
        std::cout << "[DEBUG] Using voice ID: " << WideToUtf8(voiceId) << std::endl;
    }
    
    ISpVoice* pVoice = NULL;
    hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] CoCreateInstance(CLSID_SpVoice) failed: 0x" << std::hex << hr << std::endl;
        return wavData;
    }
    std::cout << "[DEBUG] ISpVoice created successfully" << std::endl;
    
    if (!voiceId.empty()) {
        ISpObjectToken* pToken = NULL;
        hr = SpGetTokenFromId(voiceId.c_str(), &pToken, FALSE);
        if (SUCCEEDED(hr)) {
            hr = pVoice->SetVoice(pToken);
            if (SUCCEEDED(hr)) {
                std::cout << "[DEBUG] Voice set successfully" << std::endl;
            } else {
                std::cerr << "[WARNING] ISpVoice::SetVoice failed: 0x" << std::hex << hr << std::endl;
            }
            pToken->Release();
        } else {
            std::cerr << "[WARNING] SpGetTokenFromId failed: 0x" << std::hex << hr << std::endl;
        }
    }
    
    CSpStreamFormat format;
    format.AssignFormat(SPSF_22kHz16BitMono);
    std::cout << "[DEBUG] Stream format set to 22kHz 16-bit Mono" << std::endl;
    
    HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, 0);
    if (!hGlobal) {
        std::cerr << "[ERROR] GlobalAlloc failed" << std::endl;
        pVoice->Release();
        return wavData;
    }
    
    IStream* pMemStream = NULL;
    hr = CreateStreamOnHGlobal(hGlobal, TRUE, &pMemStream);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] CreateStreamOnHGlobal failed: 0x" << std::hex << hr << std::endl;
        GlobalFree(hGlobal);
        pVoice->Release();
        return wavData;
    }
    std::cout << "[DEBUG] IStream created successfully" << std::endl;
    
    ISpStream* pSpStream = NULL;
    hr = CoCreateInstance(CLSID_SpStream, NULL, CLSCTX_ALL, IID_ISpStream, (void**)&pSpStream);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] CoCreateInstance(CLSID_SpStream) failed: 0x" << std::hex << hr << std::endl;
        pMemStream->Release();
        pVoice->Release();
        return wavData;
    }
    std::cout << "[DEBUG] ISpStream created successfully" << std::endl;
    
    hr = pSpStream->SetBaseStream(pMemStream, SPDFID_WaveFormatEx, format.WaveFormatExPtr());
    if (FAILED(hr)) {
        std::cerr << "[ERROR] ISpStream::SetBaseStream failed: 0x" << std::hex << hr << std::endl;
        pSpStream->Release();
        pMemStream->Release();
        pVoice->Release();
        return wavData;
    }
    std::cout << "[DEBUG] Base stream set successfully" << std::endl;
    
    hr = pVoice->SetOutput(pSpStream, TRUE);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] ISpVoice::SetOutput failed: 0x" << std::hex << hr << std::endl;
        pSpStream->Release();
        pMemStream->Release();
        pVoice->Release();
        return wavData;
    }
    std::cout << "[DEBUG] Output set to stream" << std::endl;
    
    std::cout << "[DEBUG] Starting Speak..." << std::endl;
    hr = pVoice->Speak(text.c_str(), SPF_DEFAULT, NULL);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] ISpVoice::Speak failed: 0x" << std::hex << hr << std::endl;
        pVoice->SetOutput(NULL, FALSE);
        pSpStream->Release();
        pMemStream->Release();
        pVoice->Release();
        return wavData;
    }
    std::cout << "[DEBUG] Speak succeeded, waiting for completion..." << std::endl;
    
    hr = pVoice->WaitUntilDone(INFINITE);
    if (FAILED(hr)) {
        std::cerr << "[WARNING] WaitUntilDone failed: 0x" << std::hex << hr << std::endl;
    }
    std::cout << "[DEBUG] Speech completed" << std::endl;
    
    pVoice->SetOutput(NULL, FALSE);
    
    LARGE_INTEGER liZero = {0};
    ULARGE_INTEGER ulPos;
    hr = pMemStream->Seek(liZero, STREAM_SEEK_CUR, &ulPos);
    if (FAILED(hr)) {
        std::cerr << "[ERROR] IStream::Seek (CUR) failed: 0x" << std::hex << hr << std::endl;
    } else {
        DWORD dwSize = (DWORD)ulPos.QuadPart;
        std::cout << "[DEBUG] Stream size: " << dwSize << " bytes" << std::endl;
        
        if (dwSize > 0) {
            pMemStream->Seek(liZero, STREAM_SEEK_SET, NULL);
            
            wavData.resize(dwSize);
            ULONG ulRead = 0;
            hr = pMemStream->Read(wavData.data(), dwSize, &ulRead);
            if (FAILED(hr)) {
                std::cerr << "[ERROR] IStream::Read failed: 0x" << std::hex << hr << std::endl;
                wavData.clear();
            } else {
                std::cout << "[DEBUG] Read " << ulRead << " bytes from stream" << std::endl;
            }
        } else {
            std::cerr << "[ERROR] Stream is empty" << std::endl;
        }
    }
    
    pSpStream->Release();
    pMemStream->Release();
    pVoice->Release();
    
    std::cout << "[DEBUG] TextToWavInternal finished, WAV data size: " << wavData.size() << std::endl;
    return wavData;
}

std::vector<BYTE> ConvertWavToMp3(const std::vector<BYTE>& wavData) {
    std::vector<BYTE> mp3Data;
    
    if (!hLameDll || !lame_init) {
        std::cout << "[INFO] LAME library not loaded, skipping MP3 conversion" << std::endl;
        return mp3Data;
    }
    
    if (wavData.size() < 44) {
        std::cerr << "[ERROR] WAV data too small: " << wavData.size() << " bytes" << std::endl;
        return mp3Data;
    }
    
    const BYTE* wavHeader = wavData.data();
    int sampleRate = *(int*)(wavHeader + 24);
    int numChannels = wavHeader[22];
    int bitsPerSample = wavHeader[34];
    
    std::cout << "[DEBUG] WAV info: " << sampleRate << "Hz, " << numChannels << " channel(s), " << bitsPerSample << " bits" << std::endl;
    
    if (bitsPerSample != 16) {
        std::cerr << "[ERROR] Only 16-bit WAV format supported" << std::endl;
        return mp3Data;
    }
    
    int dataOffset = 44;
    for (int i = 12; i < (int)wavData.size() - 8; i++) {
        if (wavData[i] == 'd' && wavData[i+1] == 'a' && wavData[i+2] == 't' && wavData[i+3] == 'a') {
            dataOffset = i + 8;
            break;
        }
    }
    
    const short* pcmData = (const short*)(wavData.data() + dataOffset);
    int pcmSamples = (int)((wavData.size() - dataOffset) / 2 / numChannels);
    std::cout << "[DEBUG] PCM samples: " << pcmSamples << std::endl;
    
    auto gf = lame_init();
    if (!gf) {
        std::cerr << "[ERROR] lame_init failed" << std::endl;
        return mp3Data;
    }
    
    lame_set_num_channels(gf, numChannels);
    lame_set_in_samplerate(gf, sampleRate);
    if (lame_set_mode) {
        lame_set_mode(gf, numChannels == 1 ? 3 : 1);
    }
    
    if (lame_init_params(gf) < 0) {
        std::cerr << "[ERROR] lame_init_params failed" << std::endl;
        lame_close(gf);
        return mp3Data;
    }
    
    int mp3BufferSize = (int)(pcmSamples * 1.25 + 7200);
    mp3Data.resize(mp3BufferSize);
    
    int mp3Size;
    if (numChannels == 1) {
        mp3Size = lame_encode_buffer(gf, pcmData, NULL, pcmSamples, mp3Data.data(), mp3BufferSize);
    } else {
        mp3Size = lame_encode_buffer(gf, pcmData, pcmData + 1, pcmSamples, mp3Data.data(), mp3BufferSize);
    }
    
    if (mp3Size < 0) {
        std::cerr << "[ERROR] lame_encode_buffer failed: " << mp3Size << std::endl;
        lame_close(gf);
        mp3Data.clear();
        return mp3Data;
    }
    
    int flushSize = lame_encode_flush(gf, mp3Data.data() + mp3Size, mp3BufferSize - mp3Size);
    if (flushSize > 0) {
        mp3Size += flushSize;
    }
    
    lame_close(gf);
    
    mp3Data.resize(mp3Size);
    std::cout << "[DEBUG] MP3 size: " << mp3Size << " bytes" << std::endl;
    
    return mp3Data;
}

std::string GetHttpResponse(const std::string& content, const std::string& contentType, int statusCode = 200) {
    std::ostringstream response;
    response << "HTTP/1.1 " << statusCode << " OK\r\n";
    response << "Content-Type: " << contentType << "\r\n";
    response << "Content-Length: " << content.length() << "\r\n";
    response << "Connection: close\r\n";
    response << "Access-Control-Allow-Origin: *\r\n";
    response << "\r\n";
    response << content;
    return response.str();
}

std::string GetBinaryHttpResponse(const std::vector<BYTE>& data, const std::string& contentType, const std::string& disposition = "") {
    std::ostringstream headers;
    headers << "HTTP/1.1 200 OK\r\n";
    headers << "Content-Type: " << contentType << "\r\n";
    headers << "Content-Length: " << data.size() << "\r\n";
    headers << "Connection: close\r\n";
    headers << "Access-Control-Allow-Origin: *\r\n";
    if (!disposition.empty()) {
        headers << "Content-Disposition: " << disposition << "\r\n";
    }
    headers << "\r\n";
    
    std::string headersStr = headers.str();
    std::string response;
    response.reserve(headersStr.size() + data.size());
    response.append(headersStr);
    response.append((const char*)data.data(), data.size());
    return response;
}

std::string GetHtmlPage() {
    return R"(
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Text to Speech</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
            width: 100%;
            max-width: 600px;
        }
        h1 { text-align: center; color: #333; margin-bottom: 30px; font-size: 28px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 8px; color: #555; font-weight: 600; }
        textarea {
            width: 100%; height: 150px; padding: 15px;
            border: 2px solid #e0e0e0; border-radius: 8px;
            font-size: 16px; resize: vertical;
            transition: border-color 0.3s; font-family: inherit;
        }
        textarea:focus { outline: none; border-color: #667eea; }
        select {
            width: 100%; padding: 12px 15px;
            border: 2px solid #e0e0e0; border-radius: 8px;
            font-size: 16px; background: white;
            cursor: pointer; transition: border-color 0.3s;
        }
        select:focus { outline: none; border-color: #667eea; }
        .btn-group { display: flex; gap: 15px; margin-top: 30px; }
        button {
            flex: 1; padding: 15px; border: none; border-radius: 8px;
            font-size: 16px; font-weight: 600; cursor: pointer;
            transition: all 0.3s;
        }
        .btn-play {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        .btn-play:hover { transform: translateY(-2px); box-shadow: 0 5px 20px rgba(102, 126, 234, 0.4); }
        .btn-download {
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            color: white;
        }
        .btn-download:hover { transform: translateY(-2px); box-shadow: 0 5px 20px rgba(17, 153, 142, 0.4); }
        button:disabled { background: #ccc; cursor: not-allowed; transform: none; box-shadow: none; }
        .status {
            margin-top: 20px; padding: 15px; border-radius: 8px;
            text-align: center; display: none;
        }
        .status.loading { background: #e3f2fd; color: #1976d2; display: block; }
        .status.error { background: #ffebee; color: #c62828; display: block; }
        .status.success { background: #e8f5e9; color: #2e7d32; display: block; }
        .audio-container { margin-top: 20px; display: none; }
        audio { width: 100%; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Text to Speech</h1>
        <div class="form-group">
            <label for="voice">Select Voice:</label>
            <select id="voice">
                <option value="">Loading...</option>
            </select>
        </div>
        <div class="form-group">
            <label for="text">Enter Text:</label>
            <textarea id="text" placeholder="Enter text to convert..."></textarea>
        </div>
        <div class="btn-group">
            <button class="btn-play" id="btnPlay">Play Convert</button>
            <button class="btn-download" id="btnDownload">Download MP3</button>
        </div>
        <div class="status" id="status"></div>
        <div class="audio-container" id="audioContainer">
            <audio id="audioPlayer" controls></audio>
        </div>
    </div>
    <script>
        const textarea = document.getElementById('text');
        const voiceSelect = document.getElementById('voice');
        const btnPlay = document.getElementById('btnPlay');
        const btnDownload = document.getElementById('btnDownload');
        const statusDiv = document.getElementById('status');
        const audioContainer = document.getElementById('audioContainer');
        const audioPlayer = document.getElementById('audioPlayer');
        let currentAudioUrl = null;
        
        function showStatus(message, type) {
            statusDiv.textContent = message;
            statusDiv.className = 'status ' + type;
        }
        
        function setButtonsEnabled(enabled) {
            btnPlay.disabled = !enabled;
            btnDownload.disabled = !enabled;
        }
        
        async function loadVoices() {
            try {
                const response = await fetch('/voices');
                const voices = await response.json();
                console.log('Voices loaded:', voices);
                
                voiceSelect.innerHTML = '';
                if (voices.length === 0) {
                    const option = document.createElement('option');
                    option.value = '';
                    option.textContent = 'No voices available (check console)';
                    voiceSelect.appendChild(option);
                } else {
                    voices.forEach(voice => {
                        const option = document.createElement('option');
                        option.value = voice.id;
                        option.textContent = voice.name;
                        voiceSelect.appendChild(option);
                    });
                }
            } catch (error) {
                console.error('Failed to load voices:', error);
                voiceSelect.innerHTML = '<option value="">Load failed: ' + error.message + '</option>';
            }
        }
        
        async function convertText(format) {
            const text = textarea.value.trim();
            if (!text) {
                showStatus('Please enter text', 'error');
                return;
            }
            
            const voiceId = voiceSelect.value;
            setButtonsEnabled(false);
            showStatus('Converting...', 'loading');
            
            try {
                const formData = new URLSearchParams();
                formData.append('text', text);
                if (voiceId) {
                    formData.append('voice', voiceId);
                }
                
                const endpoint = format === 'wav' ? '/speak' : '/download';
                console.log('Sending request to:', endpoint);
                
                const response = await fetch(endpoint, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                });
                
                console.log('Response status:', response.status);
                
                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(errorText || 'Server error: ' + response.status);
                }
                
                const blob = await response.blob();
                console.log('Received blob:', blob.size, 'bytes, type:', blob.type);
                
                if (currentAudioUrl) {
                    URL.revokeObjectURL(currentAudioUrl);
                }
                
                currentAudioUrl = URL.createObjectURL(blob);
                
                if (format === 'wav') {
                    audioPlayer.src = currentAudioUrl;
                    audioContainer.style.display = 'block';
                    audioPlayer.play();
                    showStatus('Conversion successful!', 'success');
                } else {
                    const a = document.createElement('a');
                    a.href = currentAudioUrl;
                    a.download = 'tts_' + Date.now() + (blob.type === 'audio/mpeg' ? '.mp3' : '.wav');
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    showStatus('Download started!', 'success');
                }
                
            } catch (error) {
                console.error('Conversion error:', error);
                showStatus('Conversion failed: ' + error.message, 'error');
            } finally {
                setButtonsEnabled(true);
            }
        }
        
        btnPlay.addEventListener('click', () => convertText('wav'));
        btnDownload.addEventListener('click', () => convertText('mp3'));
        
        loadVoices();
    </script>
</body>
</html>
)";
}

std::string ParseRequest(const std::string& request, std::string& method, std::string& path, std::map<std::string, std::string>& headers, std::string& body) {
    std::istringstream stream(request);
    std::string line;
    
    if (std::getline(stream, line)) {
        std::istringstream lineStream(line);
        lineStream >> method >> path;
    }
    
    while (std::getline(stream, line) && line != "\r" && line != "") {
        size_t colon = line.find(':');
        if (colon != std::string::npos) {
            std::string key = line.substr(0, colon);
            std::string value = line.substr(colon + 1);
            
            size_t start = value.find_first_not_of(" \t");
            if (start != std::string::npos) {
                value = value.substr(start);
            }
            size_t end = value.find_last_not_of(" \t\r");
            if (end != std::string::npos) {
                value = value.substr(0, end + 1);
            }
            
            headers[key] = value;
        }
    }
    
    auto contentLengthIt = headers.find("Content-Length");
    if (contentLengthIt != headers.end()) {
        int length = atoi(contentLengthIt->second.c_str());
        if (length > 0) {
            char* buffer = new char[length + 1];
            stream.read(buffer, length);
            buffer[length] = '\0';
            body = buffer;
            delete[] buffer;
        }
    }
    
    return path;
}

std::map<std::string, std::string> ParseFormData(const std::string& body) {
    std::map<std::string, std::string> data;
    std::istringstream stream(body);
    std::string pair;
    
    while (std::getline(stream, pair, '&')) {
        size_t eq = pair.find('=');
        if (eq != std::string::npos) {
            std::string key = pair.substr(0, eq);
            std::string value = pair.substr(eq + 1);
            data[key] = UrlDecode(value);
        }
    }
    
    return data;
}

void HandleClient(SOCKET clientSocket) {
    std::cout << "\n[INFO] ===== New Client Request =====" << std::endl;
    
    ComInitializer comInit;
    if (!comInit.IsInitialized()) {
        std::cerr << "[ERROR] Failed to initialize COM for client thread" << std::endl;
        closesocket(clientSocket);
        return;
    }
    std::cout << "[DEBUG] COM initialized for client thread" << std::endl;
    
    char buffer[MAX_BUFFER];
    std::string request;
    int bytesReceived;
    
    while ((bytesReceived = recv(clientSocket, buffer, MAX_BUFFER - 1, 0)) > 0) {
        buffer[bytesReceived] = '\0';
        request += buffer;
        
        if (request.find("\r\n\r\n") != std::string::npos) {
            break;
        }
    }
    
    if (bytesReceived <= 0 && request.empty()) {
        std::cerr << "[ERROR] No data received from client" << std::endl;
        closesocket(clientSocket);
        return;
    }
    
    std::string method, path, body;
    std::map<std::string, std::string> headers;
    ParseRequest(request, method, path, headers, body);
    
    std::cout << "[INFO] " << method << " " << path << std::endl;
    
    std::string response;
    
    if (method == "GET" && (path == "/" || path == "/index.html")) {
        std::cout << "[DEBUG] Serving HTML page" << std::endl;
        std::string html = GetHtmlPage();
        response = GetHttpResponse(html, "text/html; charset=utf-8");
    }
    else if (method == "GET" && path == "/voices") {
        std::cout << "[DEBUG] Getting voice list" << std::endl;
        auto voices = GetVoiceListInternal();
        
        std::ostringstream json;
        json << "[";
        for (size_t i = 0; i < voices.size(); i++) {
            if (i > 0) json << ",";
            json << "{\"id\":\"" << JsonEscape(WideToUtf8(voices[i].first)) << "\",";
            json << "\"name\":\"" << JsonEscape(WideToUtf8(voices[i].second)) << "\"}";
        }
        json << "]";
        
        std::cout << "[DEBUG] Voice list JSON: " << json.str() << std::endl;
        response = GetHttpResponse(json.str(), "application/json; charset=utf-8");
    }
    else if (method == "POST" && (path == "/speak" || path == "/download")) {
        std::cout << "[DEBUG] Processing speech request" << std::endl;
        auto formData = ParseFormData(body);
        auto textIt = formData.find("text");
        
        if (textIt == formData.end() || textIt->second.empty()) {
            std::cerr << "[ERROR] No text provided" << std::endl;
            response = GetHttpResponse("Please enter text", "text/plain", 400);
        } else {
            std::wstring text = Utf8ToWide(textIt->second);
            std::wstring voiceId;
            auto voiceIt = formData.find("voice");
            if (voiceIt != formData.end()) {
                voiceId = Utf8ToWide(voiceIt->second);
            }
            
            std::cout << "[DEBUG] Converting text: " << textIt->second << std::endl;
            std::vector<BYTE> audioData = TextToWavInternal(text, voiceId);
            
            if (audioData.empty()) {
                std::cerr << "[ERROR] TextToWavInternal returned empty data" << std::endl;
                response = GetHttpResponse("Speech conversion failed - check server logs for details", "text/plain", 500);
            } else {
                std::cout << "[INFO] WAV generated successfully, size: " << audioData.size() << " bytes" << std::endl;
                
                if (path == "/download") {
                    std::vector<BYTE> mp3Data = ConvertWavToMp3(audioData);
                    if (!mp3Data.empty()) {
                        std::cout << "[INFO] MP3 conversion successful" << std::endl;
                        response = GetBinaryHttpResponse(mp3Data, "audio/mpeg", "attachment; filename=tts.mp3");
                    } else {
                        std::cout << "[INFO] Falling back to WAV for download" << std::endl;
                        response = GetBinaryHttpResponse(audioData, "audio/wav", "attachment; filename=tts.wav");
                    }
                } else {
                    response = GetBinaryHttpResponse(audioData, "audio/wav");
                }
            }
        }
    }
    else {
        std::cout << "[DEBUG] 404 Not Found: " << path << std::endl;
        response = GetHttpResponse("404 Not Found", "text/plain", 404);
    }
    
    std::cout << "[DEBUG] Sending response, size: " << response.length() << " bytes" << std::endl;
    send(clientSocket, response.c_str(), (int)response.length(), 0);
    closesocket(clientSocket);
    
    std::cout << "[INFO] ===== Request Complete =====" << std::endl;
}

int main() {
    SetConsoleOutputCP(CP_UTF8);
    
    std::cout << "========================================" << std::endl;
    std::cout << "Text-to-Speech Server" << std::endl;
    std::cout << "========================================" << std::endl;
    
    std::cout << "\n[INFO] Initializing main thread COM..." << std::endl;
    ComInitializer mainComInit;
    if (!mainComInit.IsInitialized()) {
        std::cerr << "[ERROR] Failed to initialize main thread COM" << std::endl;
        return 1;
    }
    
    std::cout << "\n[INFO] Loading LAME encoder..." << std::endl;
    LoadLameLibrary();
    
    std::cout << "\n[INFO] Testing SAPI in main thread..." << std::endl;
    {
        ComInitializer testCom;
        if (testCom.IsInitialized()) {
            auto testVoices = GetVoiceListInternal();
            std::cout << "[INFO] Main thread test - Voices available: " << testVoices.size() << std::endl;
            for (size_t i = 0; i < testVoices.size(); i++) {
                std::cout << "  [" << i << "] " << WideToUtf8(testVoices[i].second) << std::endl;
            }
        }
    }
    
    std::cout << "\n[INFO] Initializing Winsock..." << std::endl;
    WSADATA wsaData;
    if (WSAStartup(MAKEWORD(2, 2), &wsaData) != 0) {
        std::cerr << "[ERROR] WSAStartup failed" << std::endl;
        return 1;
    }
    
    SOCKET serverSocket = socket(AF_INET, SOCK_STREAM, 0);
    if (serverSocket == INVALID_SOCKET) {
        std::cerr << "[ERROR] socket creation failed" << std::endl;
        WSACleanup();
        return 1;
    }
    
    int optval = 1;
    setsockopt(serverSocket, SOL_SOCKET, SO_REUSEADDR, (const char*)&optval, sizeof(optval));
    
    sockaddr_in serverAddr;
    serverAddr.sin_family = AF_INET;
    serverAddr.sin_addr.s_addr = INADDR_ANY;
    serverAddr.sin_port = htons(PORT);
    
    if (bind(serverSocket, (sockaddr*)&serverAddr, sizeof(serverAddr)) == SOCKET_ERROR) {
        std::cerr << "[ERROR] bind failed" << std::endl;
        closesocket(serverSocket);
        WSACleanup();
        return 1;
    }
    
    if (listen(serverSocket, 10) == SOCKET_ERROR) {
        std::cerr << "[ERROR] listen failed" << std::endl;
        closesocket(serverSocket);
        WSACleanup();
        return 1;
    }
    
    std::cout << "\n========================================" << std::endl;
    std::cout << "Server Started Successfully!" << std::endl;
    std::cout << "Listening on port: " << PORT << std::endl;
    std::cout << "Please visit: http://localhost:" << PORT << std::endl;
    std::cout << "========================================" << std::endl;
    
    while (true) {
        sockaddr_in clientAddr;
        int clientLen = sizeof(clientAddr);
        SOCKET clientSocket = accept(serverSocket, (sockaddr*)&clientAddr, &clientLen);
        
        if (clientSocket == INVALID_SOCKET) {
            std::cerr << "[ERROR] accept failed" << std::endl;
            continue;
        }
        
        char clientIP[INET_ADDRSTRLEN];
        inet_ntop(AF_INET, &clientAddr.sin_addr, clientIP, INET_ADDRSTRLEN);
        std::cout << "\n[INFO] New connection from: " << clientIP << ":" << ntohs(clientAddr.sin_port) << std::endl;
        
        std::thread t(HandleClient, clientSocket);
        t.detach();
    }
    
    closesocket(serverSocket);
    WSACleanup();
    UnloadLameLibrary();
    
    return 0;
}
