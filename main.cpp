/*
 * ============================================================================
 *  PPTX 智能打开器 - selecter.exe (最优化版)
 * ============================================================================
 *  功能：根据 PPTX 文件的创建软件，自动选择 PowerPoint 或 WPS 打开
 *  场景：教室电脑，不同老师用不同软件制作 PPT，避免排版错乱
 *  技术：miniz 直接读取 PPTX（ZIP）内的 docProps/app.xml，提取 <Application>
 *  优化点：
 *    1. 零 iostream 依赖（静默模式，无控制台窗口）
 *    2. MZ_ZIP_FLAG_DO_NOT_SORT_CENTRAL_DIRECTORY 跳过不必要的排序
 *    3. QuickGetTag 使用纯 C strstr，不构造临时 std::string 标签
 *    4. 决策逻辑合并为单次 if-else，避免重复 find()
 *    5. RAII 风格的 ZipGuard 确保 zip 套柄始终被关闭
 *    6. SW_SHOWNORMAL 替代 SW_SHOW，行为更稳定
 *    7. 日志记录，便于排查问题
 * ============================================================================
 */

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <cstdio>
#include <cstring>
#include <string>

#define MINIZ_NO_TIME
#define MINIZ_NO_ARCHIVE_WRITING_APIS // 只需要读，禁用写接口
#define MINIZ_IMPLEMENTATION
#include "miniz.h"

// ============================================================================
//  日志配置
// ============================================================================
static const char* LOG_DIR = "d:\\logs";
static const char* LOG_FILE = "d:\\logs\\pptx_selecter.log";
static const bool ENABLE_LOG = true;

// ============================================================================
//  初始化日志目录：确保 d:\logs 文件夹存在
// ============================================================================
static void InitLogDir() {
    // 创建日志目录（如果不存在）
    CreateDirectoryA(LOG_DIR, nullptr);
}

// ============================================================================
//  日志函数：追加写入日志文件
// ============================================================================
static void Log(const char* format, ...) {
    if (!ENABLE_LOG) return;

    FILE* fp = nullptr;
    fopen_s(&fp, LOG_FILE, "a");
    if (!fp) return;

    // 获取当前时间
    SYSTEMTIME st;
    GetLocalTime(&st);
    fprintf(fp, "[%04d-%02d-%02d %02d:%02d:%02d] ",
            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);

    va_list args;
    va_start(args, format);
    vfprintf(fp, format, args);
    va_end(args);

    fprintf(fp, "\n");
    fclose(fp);
}

// ============================================================================
//  配置区域：根据教室电脑实际路径修改
// ============================================================================
static const char* PATH_POWERPOINT =
    "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE";

static const char* PATH_WPS =
    "C:\\Program Files (x86)\\Kingsoft\\WPS Office\\12.1.0.23542\\office6\\wpp.exe";

// ============================================================================
//  RAII ZIP 守护：确保 mz_zip_reader_end 一定会被调用
// ============================================================================
struct ZipGuard {
    mz_zip_archive& z;
    bool valid;
    explicit ZipGuard(mz_zip_archive& z) : z(z), valid(false) {}
    ~ZipGuard() { if (valid) mz_zip_reader_end(&z); }
};

// ============================================================================
//  QuickGetTag
//  纯 C 字符串操作，在 [src, src+len) 内提取 <tag>...</tag> 之间的内容。
//  不构造临时 std::string 标签，减少堆分配。
// ============================================================================
static std::string QuickGetTag(const char* src, size_t len, const char* tag) {
    if (!src || len == 0 || !tag) return {};

    // 在栈上构造 "<tag>" 和 "</tag>"（最大支持 63 字节标签名）
    char sTag[128], eTag[128];
    const int sLen = _snprintf_s(sTag, sizeof(sTag), _TRUNCATE, "<%s>", tag);
    const int eLen = _snprintf_s(eTag, sizeof(eTag), _TRUNCATE, "</%s>", tag);
    if (sLen <= 0 || eLen <= 0) return {};

    // strstr 即可：miniz 解压出的 XML 缓冲区末尾有 \0
    const char* p = strstr(src, sTag);
    if (!p) return {};
    p += sLen;

    const char* e = strstr(p, eTag);
    if (!e) return {};

    return std::string(p, static_cast<size_t>(e - p));
}

// ============================================================================
//  SafeLaunchApp
//  用 ShellExecuteA 异步启动程序，文件路径自动加引号。
//  返回：true = 启动成功（> 32），false = 失败
// ============================================================================
static bool SafeLaunchApp(const char* exe, const char* filePath) {
    // 手动拼接 "\"<filePath>\""，避免 std::string 构造开销
    char quoted[MAX_PATH + 4];
    _snprintf_s(quoted, sizeof(quoted), _TRUNCATE, "\"%s\"", filePath);

    const auto r = reinterpret_cast<intptr_t>(
        ShellExecuteA(nullptr, "open", exe, quoted, nullptr, SW_SHOWNORMAL));

    bool success = (r > 32);
    Log("启动程序: exe=%s, file=%s, result=%s (code=%d)",
        exe, filePath, success ? "成功" : "失败", (int)r);
    return success;
}

// ============================================================================
//  WinMain（无控制台窗口）
//  如需调试时看输出，将入口换回 main 并去掉 MINIZ_NO_STDIO
// ============================================================================
int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR lpCmdLine, int) {
    // 初始化日志目录
    InitLogDir();

    Log("========== 程序启动 ==========");

    // ── 1. 解析命令行，取第一个参数作为 PPTX 路径 ──────────────────────────
    int argc = 0;
    LPWSTR* argvW = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (!argvW || argc < 2) {
        Log("错误：命令行参数不足 (argc=%d)", argc);
        if (argvW) LocalFree(argvW);
        return 1;
    }

    // 宽字符路径转 ANSI（教室环境通常是 GBK/中文路径）
    char pptxPath[MAX_PATH] = {};
    WideCharToMultiByte(CP_ACP, 0, argvW[1], -1, pptxPath, MAX_PATH, nullptr, nullptr);
    LocalFree(argvW);

    if (pptxPath[0] == '\0') {
        Log("错误：路径转换失败");
        return 1;
    }

    Log("目标文件: %s", pptxPath);
    Log("PowerPoint 路径: %s", PATH_POWERPOINT);
    Log("WPS 路径: %s", PATH_WPS);

    // 检查 PowerPoint 是否存在
    bool pptExists = (GetFileAttributesA(PATH_POWERPOINT) != INVALID_FILE_ATTRIBUTES);
    Log("PowerPoint 文件存在: %s", pptExists ? "是" : "否");

    // 检查 WPS 是否存在
    bool wpsExists = (GetFileAttributesA(PATH_WPS) != INVALID_FILE_ATTRIBUTES);
    Log("WPS 文件存在: %s", wpsExists ? "是" : "否");

    // ── 2. 文件存在性检查 ────────────────────────────────────────────────────
    if (GetFileAttributesA(pptxPath) == INVALID_FILE_ATTRIBUTES) {
        Log("错误：文件不存在，尝试用 WPS 兜底打开");
        SafeLaunchApp(PATH_WPS, pptxPath); // 兜底
        return 1;
    }

    // ── 3. 打开 ZIP（PPTX 本质是 ZIP） ──────────────────────────────────────
    mz_zip_archive zip = {};
    ZipGuard guard(zip);

    if (!mz_zip_reader_init_file(&zip, pptxPath,
                                  MZ_ZIP_FLAG_DO_NOT_SORT_CENTRAL_DIRECTORY)) {
        Log("错误：无法打开 ZIP 文件，尝试用 WPS 兜底打开");
        SafeLaunchApp(PATH_WPS, pptxPath);
        return 1;
    }
    guard.valid = true;

    // ── 4. 定位并解压 docProps/app.xml ──────────────────────────────────────
    std::string appName;
    const int idx = mz_zip_reader_locate_file(&zip, "docProps/app.xml", nullptr, 0);
    if (idx >= 0) {
        size_t sz = 0;
        char* pXml = static_cast<char*>(
            mz_zip_reader_extract_to_heap(&zip, idx, &sz, 0));
        if (pXml) {
            appName = QuickGetTag(pXml, sz, "Application");
            mz_free(pXml);
        }
    }
    // guard 析构时自动调用 mz_zip_reader_end

    Log("检测到 Application: %s", appName.empty() ? "(空)" : appName.c_str());

    // ── 5. 决策：优先精确匹配，再宽泛匹配，最后兜底 ─────────────────────────
    bool launched = false;

    if (!appName.empty()) {
        const bool isWPS  = (appName.find("WPS")      != std::string::npos)
                          || (appName.find("Kingsoft") != std::string::npos);
        const bool isMSO  = (appName.find("Microsoft") != std::string::npos)
                          || (appName.find("PowerPoint") != std::string::npos);

        Log("判断结果: isWPS=%d, isMSO=%d", isWPS, isMSO);

        if (isWPS) {
            Log("决策：使用 WPS 打开");
            launched = SafeLaunchApp(PATH_WPS, pptxPath);
        } else if (isMSO) {
            Log("决策：使用 PowerPoint 打开");
            launched = SafeLaunchApp(PATH_POWERPOINT, pptxPath);
        }
    }

    if (!launched) {
        Log("兜底：使用 WPS 打开");
        SafeLaunchApp(PATH_WPS, pptxPath); // 兜底 WPS
    }

    Log("========== 程序结束 ==========\n");
    return 0;
}
