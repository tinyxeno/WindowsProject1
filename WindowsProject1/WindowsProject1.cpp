// WindowsProject1.cpp : アプリケーションのエントリ ポイントを定義します。
//

#include "framework.h"
#include "WindowsProject1.h"
#pragma comment(lib,"comctl32.lib")
#pragma comment(lib,"imm32.lib")
#pragma comment(lib,"htmlhelp.lib")
#if _MSC_VER < 1600
#pragma comment(lib,"version.lib")
#else
#ifdef _UNICODE
#pragma comment(lib,"Mincore.lib")
#else
#pragma comment(lib,"version.lib")
#endif
#endif
#define NULLSTR _T('\0')
#define DEF_PRINTER

/// <summary>
/// 連想リスト
/// </summary>
typedef struct _tagASSOC
{
    UINT key;
    LPCTSTR value;
}ASSOC, * LPASSOC;

/// <summary>
/// リストビューのソート用構造体
/// </summary>
typedef struct _tagSORTLIST
{
    HWND hWnd;
    int iSubItem;
    BOOL bDirection;
}SORTLIST, * LPSORTLIST;

/// <summary>
/// ツールバーリソースの取得用構造体
/// </summary>
typedef struct _tagTOOLBARDATA
{
    WORD version;
    WORD width;
    WORD height;
    WORD count;
}TOOLBARDATA, * LPTOOLBARDATA;

// グローバル変数:
HINSTANCE m_hInst;        // 現在のインターフェイス
LPCTSTR m_szTitle;        // タイトル バーのテキスト
LPCTSTR m_szWindowClass;  // メイン ウィンドウ クラス名

/// <summary>
/// ステータスバーへ表示するための現在日時
/// </summary>
SYSTEMTIME  m_systemTime;

/// <summary>
/// 「ファイルを開く」や「名前をつけて保存」で使用する構造体
/// </summary>
LPOPENFILENAME  m_pOpenFileName;

/// <summary>
/// 印刷処理構造体
/// </summary>
LPPRINTDLG      m_pPrintDlg;

/// <summary>
/// ページ設定処理構造体
/// </summary>
LPPAGESETUPDLG  m_pPageSetupDlg;

/// <summary>
/// 更新フラグ
/// </summary>
BOOL    m_bDirty;

/// <summary>
/// メイン・ウィンドウのハンドル
/// </summary>
HWND    m_hMainWnd;

/// <summary>
/// 印刷プレビュー中のメニューバーの保持先
/// </summary>
HMENU   m_hMenu;

/// <summary>
/// ツールバーのツールチップ表示用の文字列
/// </summary>
LPTSTR  m_sToolTip;

/// <summary>
/// ステータス・バーのコマンドIDの配列
/// </summary>
UINT*   m_pStatusBar;

/// <summary>
/// ステータス・バーへの表示文字列の配列
/// </summary>
LPCTSTR* m_aStatusBar;

/// <summary>
/// ステータス・バーの許可/禁止フラグの配列
/// </summary>
BOOL*   m_bStatusBar;

/// <summary>
/// リスト・ビューのサブクラスアドレスの保持先
/// </summary>
WNDPROC m_wndList;

/// <summary>
/// リスト・ビューで注目している現在位置
/// </summary>
int     m_iItem, m_iSubItem;

/// <summary>
/// 検索・置換ダイアログのウィンドウハンドル
/// </summary>
HWND    m_hFindReplace;

/// <summary>
/// 検索・置換ダイアログの処理メッセージID
/// </summary>
UINT    WM_FINDREPLACE;

/// <summary>
/// 検索・置換ダイアログ用構造体
/// </summary>
LPFINDREPLACE   m_pFindReplace;

/// <summary>
/// ヘルプウィンドウのハンドル
/// </summary>
/// <param name="hInstance"></param>
/// <returns></returns>
HWND m_hHelp;

/// <summary>
/// レコードセット
/// </summary>
IDispatch* m_adoRS;

/// <summary>
/// フィールド・コレクション
/// </summary>
IDispatch* m_adoFDS;

/// <summary>
/// プリンタへのハンドル(中断処理用)
/// </summary>
HANDLE  m_hPrinter;

/// <summary>
/// プリンタ名
/// </summary>
LPCTSTR m_pPrinterName;

/// <summary>
/// 印刷中のページ番号
/// </summary>
int     m_nPrintCurrentPage;

/// <summary>
/// 最大印刷ページ番号
/// </summary>
int     m_nPrintMaxPage;

/// <summary>
/// 印刷フィールドの列幅
/// </summary>
int     m_nPrintFieldPitch;

/// <summary>
/// 印刷の行高さ
/// </summary>
int     m_nPrintRowPitch;

/// <summary>
/// １ページ当たりの行数
/// </summary>
int     m_nPrintLineByPage;

/// <summary>
/// プレビュー表示ページ数
/// </summary>
int     m_nPreviewpageNum;

/// <summary>
/// 左マージン
/// </summary>
int     m_nPrintMarginLeft;

/// <summary>
/// 上マージン
/// </summary>
int     m_nPrintMarginTop;

// このコード モジュールに含まれる関数の宣言を転送します:

/// <summary>
/// ウィンドウ クラスを登録します。
/// </summary>
/// <param name="hInstance">インスタンスハンドル</param>
/// <returns>登録状態</returns>
ATOM                MyRegisterClass(HINSTANCE hInstance);

/// <summary>
/// インスタンス ハンドルを保存して、メイン ウィンドウを作成します
/// </summary>
/// <param name="">インスタンスハンドル</param>
/// <param name="">表示開始フラグ</param>
/// <returns>メインウィンドウの処理が成功した場合は偽以外を返す</returns>
BOOL                InitInstance(HINSTANCE, int);

/// <summary>
/// メイン ウィンドウのメッセージを処理します。
/// </summary>
/// <param name=""></param>
/// <param name=""></param>
/// <param name=""></param>
/// <param name=""></param>
/// <returns></returns>
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);

/// <summary>
/// バージョン情報ボックスのメッセージ ハンドラーです。
/// </summary>
/// <param name="">メッセージボックスダイアログのハンドル</param>
/// <param name="">メッセージID</param>
/// <param name="">パラメータ</param>
/// <param name="">パラメータ</param>
/// <returns></returns>
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

/// <summary>
/// 印刷設定ダイアログのコールバック
/// (特に何もしていない。設定しないと、デバイス依存のダイアログが起動する)
/// </summary>
/// <param name=""></param>
/// <param name="message"></param>
/// <param name="wParam"></param>
/// <param name="lParam"></param>
/// <returns></returns>
UINT_PTR CALLBACK   PrintSetupProc(HWND, UINT, WPARAM, LPARAM);

/// <summary>
/// 印刷中ダイアログボックスのメッセージ ハンドラーです。
/// </summary>
/// <param name="">メッセージボックスダイアログのハンドル</param>
/// <param name="">メッセージID</param>
/// <param name="">パラメータ</param>
/// <param name="">パラメータ</param>
/// <returns></returns>
INT_PTR CALLBACK    PrintProc(HWND, UINT, WPARAM, LPARAM);

/// <summary>
/// リストビューの処理
/// </summary>
/// <param name="">リストビューのハンドル</param>
/// <param name="">メッセージID</param>
/// <param name="">パラメータ</param>
/// <param name="">パラメータ</param>
/// <returns></returns>
LRESULT CALLBACK    ListProc(HWND, UINT, WPARAM, LPARAM);

/// <summary>
/// リストビューのソート判定関数
/// </summary>
/// <param name="lParam1">要素番号１</param>
/// <param name="lParam2">要素番号２</param>
/// <param name="lParam3">ソートの緒元</param>
/// <returns>大小を返す。等しい場合は0</returns>
int     CALLBACK    SortList(LPARAM lParam1, LPARAM lParam2, LPARAM lParam3);

/// <summary>
/// 文字列を指定文字列で分割し、配列へ変換して返す
/// </summary>
/// <param name="results">分割後の配列</param>
/// <param name="source">検索対象文字列</param>
/// <param name="delimiter">分割子</param>
/// <returns>分割した個数</returns>
UINT Split(LPCTSTR** results, LPCTSTR source = nullptr, 
    LPCTSTR delimiter = _T("\n"));

/// <summary>
/// 文字列の置き換え
/// </summary>
/// <param name="source">検索対象文字列</param>
/// <param name="pOld">検索文字列</param>
/// <param name="pNew">置換文字列</param>
/// <param name="pResult">置換個数を返すポインタ</param>
/// <returns>置換後の文字列。置換できなければ、検索文字列をコピーして返す
/// </returns>
LPCTSTR replace(LPCTSTR source, LPCTSTR pOld, LPCTSTR pNew,
    UINT* pResult = nullptr);

/// <summary>
/// 文字列を指定文字列で分割し、指定位置の文字列を返す
/// </summary>
/// <param name="source">検索対象文字列</param>
/// <param name="position">分割位置</param>
/// <param name="delimiter">分割子</param>
/// <returns></returns>
LPCTSTR AfxExtractSubString(LPCTSTR source, int position,
    LPCTSTR delimiter = _T("\n"));

/// <summary>
/// 文字列をフォーマットして返す
/// </summary>
/// <param name="nID">文字列リソースのID</param>
/// <param name="arg1"></param>
/// <param name="arg2"></param>
/// <returns></returns>
LPCTSTR AfxFormatString(UINT nID, LPCTSTR arg1, LPCTSTR arg2 = nullptr);

/// <summary>
/// メッセージボックス（ID版）
/// </summary>
/// <param name="nID">文字列ID</param>
/// <param name="style">メッセージボックスの表示型</param>
/// <returns>メッセージボックスの操作結果</returns>
UINT    MessageBox(UINT nID, UINT style = MB_ICONEXCLAMATION);

/// <summary>
/// メッセージボックス（文字列版）
/// </summary>
/// <param name="nID">文字列</param>
/// <param name="style">メッセージボックスの表示型</param>
/// <returns>メッセージボックスの操作結果</returns>
UINT    MessageBox(LPCTSTR caption, UINT style = MB_ICONEXCLAMATION);

/// <summary>
/// ツールバーのアイコンを返す
/// </summary>
/// <param name="nID">コマンドID</param>
/// <returns>コマンドIDに該当するアイコンうぃ返す</returns>
HICON   GetToolBarIcon(UINT nID);

/// <summary>
/// ステータスバーの列番号をコマンド番号から検索する
/// </summary>
/// <param name="nID">コマンド番号</param>
/// <returns>列位置。発見できなければ-1を返す</returns>
int     CommandToIndex(UINT nID);

/// <summary>
/// ドキュメントの保存を確認し、必要であれば保存する
/// </summary>
/// <returns>ドキュメントが破棄加納であれば偽以外を返す</returns>
BOOL    SaveModified();

/// <summary>
/// リストビューの列数を返す
/// </summary>
/// <returns></returns>
int     GetSubItemCount();

/// <summary>
/// リストビューアイテムの読み書きを行う
/// </summary>
/// <param name="hWnd"></param>
/// <param name="iSubItem"></param>
/// <param name="source"></param>
/// <returns></returns>
LPCTSTR DDX_ListViewItemText(HWND hWnd, 
    int iSubItem, LPCTSTR source = nullptr);

/// <summary>
/// リストビューアイテムにフォーカスを設定する
/// </summary>
/// <param name="iSubItem"></param>
void SetRowCol(HWND hWnd);

/// <summary>
/// 検索条件に従い、リストビューのアイテムを確認する
/// </summary>
/// <param name="hWnd">リストビューのハンドル</param>
/// <param name="iSubItem">リストビューの列位置</param>
/// <returns>条件に合致する場合、偽以外を返す</returns>
BOOL FindItem(HWND hWnd, int iSubItem);

/// <summary>
/// 文字列を BSTR へ変換します。
/// </summary>
/// <param name="pValue">文字列</param>
/// <returns>BSTR文字列</returns>
BSTR    str2BSTR(LPCTSTR	pValue);

/// <summary>
/// BSTR を文字列へ変換します。
/// </summary>
/// <param name="bstr">BSTR文字列</param>
/// <returns>文字列</returns>
LPCTSTR	BSTR2str(BSTR	bstr);

/// <summary>
/// COM オブジェクトを作成します。
/// </summary>
/// <param name="ppDispatch">COMオブジェクト</param>
/// <param name="pClassName">オブジェクト名</param>
/// <returns>処理結果</returns>
HRESULT	create_instance(IDispatch** ppDispatch, LPCTSTR		pClassName);

/// <summary>
/// COMオブジェクトを操作します。
/// </summary>
/// <param name="ppDispatch">COM オブジェクト</param>
/// <param name="pCommand">コマンド文字列</param>
/// <param name="bPut">【省略可】Get/Putフラグ。省略時、Put</param>
/// <param name="pResult">【省略可】戻り値</param>
/// <param name="pParam">【省略可】パラメータの配列</param>
/// <param name="uParamCount">【省略可】パラメータの個数。省略時１個</param>
/// <returns>処理結果</returns>
HRESULT	com_invoke(
    IDispatch   _In_*       source,
    LPCTSTR     _In_        pCommand,
    BOOL        _In_        bPut = TRUE,
    VARIANT     _Out_opt_*  pResult = nullptr,
    VARIANT     _In_opt_*   pParam = nullptr,
    UINT        _In_        uParamCount = 1L);

/// <summary>
/// DBの接続文字列を作成します。
/// </summary>
/// <returns>DBの接続文字列</returns>
LPCTSTR GetDefaultConnect();

/// <summary>
/// スキーマー文字列を作成します。
/// </summary>
/// <returns></returns>
LPCTSTR GetDefaultSQL();

/// <summary>
/// DBをオープンします。
/// </summary>
/// <returns>成否。成功の場合、偽以外を返す</returns>
BOOL    OpenDB();

/// <summary>
/// DBを閉じる
/// </summary>
void    CloseDB();

/// <summary>
/// DBがオープンされているかどうかを判定する
/// </summary>
/// <returns>データベースがオープンされている場合、偽以外を返す</returns>
BOOL    IsOpen();

/// <summary>
/// EOF状態を返す
/// </summary>
/// <returns>EOFの場合、偽以外を返す</returns>
BOOL    IsEOF();

/// <summary>
/// 次のレコードへ移動
/// </summary>
void    MoveNext();

/// <summary>
/// フィールドの個数を返す
/// </summary>
/// <returns>フィールドの個数</returns>
long    GetFieldCount();

/// <summary>
/// フィールド名を返す
/// </summary>
/// <param name="index">フィールド位置</param>
/// <returns>フィールド名</returns>
LPCTSTR GetFieldName(long index);

/// <summary>
/// フィールドの値を返す
/// </summary>
/// <param name="index">フィールド位置</param>
/// <returns>フィールドの値</returns>
LPCTSTR GetFieldValue(long index);

/// <summary>
/// 印刷設定
/// </summary>
/// <param name="hDevmode">プリンタデバイスへのハンドル</param>
/// <returns>エラーメッセージ。エラーでない場合、nullptr を返す</returns>
LPCTSTR SetPrinter(HGLOBAL hDevMode, HWND hWnd = nullptr);

/// <summary>
/// 印刷処理
/// </summary>
/// <param name="hDevmode">プリンタデバイスへのハンドル</param>
/// <returns>エラーメッセージ。エラーでない場合、nullptr を返す</returns>
/// <returns>処理結果。処理成功の場合、偽以外を返す</returns>
BOOL OnPrint(HDC hDC, int nPage);

//■ 以下、定数

/// <summary>
/// 起動時のウィンドウサイズ
/// </summary>
SIZE    m_WindowStartUpSize = { 833,341 };

/// <summary>
/// 最小ウィンドウサイズ
/// </summary>
SIZE    m_WindowMinSize = { 433,170 };

/// <summary>
/// アイコン定数
/// </summary>
UINT    m_cuSystemIcons[] = { 106,103,102,101,104 };

/// <summary>
/// ファイルのオープン時に許可/禁止が設定されるコマンドIDの一覧
/// </summary>
UINT    m_cuEnables[] =
{
    ID_FILE_NEW,
    ID_FILE_OPEN,
    ID_FILE_SAVE,
    ID_FILE_SAVE_AS,
    ID_FILE_PAGE_SETUP,
    ID_FILE_PRINT_PREVIEW,
    ID_FILE_PRINT_DIRECT,
    ID_FILE_PRINT,
    ID_FILE_SEND_MAIL,
    ID_EDIT_FIND,
    ID_EDIT_REPEAT,
    ID_EDIT_REPLACE,
    ID_EDIT_CLEAR,
    ID_OLE_EDIT_PROPERTIES,
    ID_OLE_INSERT_NEW,
};

/// <summary>
/// タイマー処理で許可禁止されるコマンドIDの一覧
/// </summary>
UINT    m_cuEnablesTimer[] =
{
    ID_EDIT_COPY,
    ID_EDIT_CUT,
    ID_EDIT_PASTE,
    ID_EDIT_SELECT_ALL,
    ID_EDIT_UNDO,
};

/// <summary>
/// リストビューオブジェクトの表示変更コマンドの一覧
/// </summary>
UINT    m_cuListViewCommands[] =
{
    ID_VIEW_SMALLICON,
    ID_VIEW_LARGEICON,
    ID_VIEW_LIST,
    ID_VIEW_DETAILS,
};

/// <summary>
/// 印刷プレビューのコマンドボタン一覧
/// </summary>
UINT    m_cuPreviewBar[] =
{
    AFX_ID_PREVIEW_PRINT,
    AFX_ID_PREVIEW_PREV,
    AFX_ID_PREVIEW_NEXT,
    AFX_ID_PREVIEW_NUMPAGE,
    AFX_ID_PREVIEW_ZOOMIN,
    AFX_ID_PREVIEW_ZOOMOUT,
    AFX_ID_PREVIEW_CLOSE,
};

/// <summary>
/// 印刷プレビュー時に表示されるツールバーのアイコン
/// </summary>
UINT    m_cuPreviewBarIcons[] =
{
    MAKELONG(AFX_ID_PREVIEW_PRINT,      MAKEWORD(STD_PRINT,         IDB_STD_SMALL_COLOR)),
    MAKELONG(AFX_ID_PREVIEW_PREV,       MAKEWORD(HIST_BACK,         IDB_HIST_SMALL_COLOR)),
    MAKELONG(AFX_ID_PREVIEW_NEXT,       MAKEWORD(HIST_FORWARD,      IDB_HIST_SMALL_COLOR)),
    MAKELONG(AFX_ID_PREVIEW_NUMPAGE,    MAKEWORD(HIST_FAVORITES,    IDB_HIST_SMALL_COLOR)),
    MAKELONG(AFX_ID_PREVIEW_ZOOMIN,     MAKEWORD(STD_FIND,          IDB_STD_SMALL_COLOR)),
    MAKELONG(AFX_ID_PREVIEW_ZOOMOUT,    MAKEWORD(STD_REPLACE,       IDB_STD_SMALL_COLOR)),
    MAKELONG(AFX_ID_PREVIEW_CLOSE,      MAKEWORD(STD_DELETE,        IDB_STD_SMALL_COLOR)),
};

/// <summary>
/// アプリケーション情報の設定用文字列の取得リスト
/// </summary>
LPCTSTR m_pAboutValueKeys[] =
{
    _T("ProductName"),
    _T("ProductVersion"),
    _T("LegalCopyright"),
    _T("CompanyName"),
    _T("FileDescription"),
};

/// <summary>
/// アクセラレータキーのプリフィックス
/// </summary>
ASSOC   m_aControlKey[] =
{
    {   FCONTROL,   _T("Ctrl+") },
    {   FSHIFT,     _T("Shift+")},
    {   FALT,       _T("Alt+")  },
};

/// <summary>
/// アクセラレータキーの固定キーの名称
/// </summary>
ASSOC   m_aValueKey[] =
{
    {   VK_INSERT,  _T("Ins")   },
    {   VK_DELETE,  _T("Del")   },
    {   VK_RETURN,  _T("Enter") },
    {   VK_BACK,    _T("BS")    },
    {   _T('?'),    _T("?")     },
    {   _T('/'),    _T("?")     },
    {   VK_OEM_2,   _T("?")     },
};

int APIENTRY _tWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPTSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: ここにコードを挿入してください。
    int result = 1;
    // アプリケーション初期化の実行:
    void* value = ::_tsetlocale(LC_ALL, _T("japanese"));
    if (nullptr != value &&
        SUCCEEDED(::CoInitialize(0)))
    {
        ::InitCommonControls();
        if (::MyRegisterClass(hInstance) &&
            ::InitInstance(hInstance, nCmdShow) &&
            ::IsWindow(::m_hMainWnd))
        {
            // リソースからアクセラレータを取得
            HACCEL source = ::LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_MAINFRAME));
            if (nullptr != source &&
                INVALID_HANDLE_VALUE != source)
            {
                // アクセラレータ情報をメニューバーへ反映する
                // メニューバーのハンドルを取得
                HMENU hMenu = ::GetMenu(::m_hMainWnd);
                if (::IsMenu(hMenu))
                {
                    // アクセラレータ情報の個数を取得する
                    int value = ::CopyAcceleratorTable(source, nullptr, 0u);
                    if (0 < value)
                    {
                        // アクセラレータ情報取得用配列を構築する
                        int index = value * sizeof(ACCEL);
                        LPACCEL result = (LPACCEL)::malloc(index);
                        if (nullptr != result)
                        {
                            // アクセラレータ情報を取得する
                            if (value == ::CopyAcceleratorTable(source, result, value))
                            {
                                LPACCEL source = result;
                                // メニュー情報取得構造体を作成する
                                index = sizeof(MENUITEMINFO);
                                LPMENUITEMINFO result = (LPMENUITEMINFO)::malloc(index);
                                if (nullptr != result)
                                {
                                    ::SecureZeroMemory(result, index);
                                    result->cbSize      = (UINT)index;
                                    result->fMask       = MIIM_STRING;
                                    result->dwTypeData  = (LPTSTR)::m_szWindowClass;

                                    // アクセラレータ情報の個数でループ
                                    for (int index = 0; index < value; index++)
                                    {
                                        // アクセラレータ情報を取得する
                                        LPACCEL value = &source[index];
                                        {
                                            LPACCEL source = value;

                                            // アクセラレータIDに該当するメニューアイテムに設定されている文字列を取得する
                                            int length = 1 + MAX_PATH;
                                            size_t value = sizeof(TCHAR) * length;
                                            ::SecureZeroMemory(result->dwTypeData, value);
                                            if (0 < ::GetMenuString(hMenu, source->cmd, result->dwTypeData, length, MF_BYCOMMAND))
                                            {
                                                // アクセラレータ文字列がメニューバーアイテムに既に設定されているかどうかを判定する（先勝ち１個のみ）
                                                LPCTSTR value = ::AfxExtractSubString(result->dwTypeData, 1, _T("\t"));
                                                if (nullptr == value)
                                                {
                                                    // メニューアイテム文字列にアクセラレータ文字列が設定されていない場合
                                                    ::_tcscat_s(result->dwTypeData, length, _T("\t"));
                                                    // アクセラレータ文字列のプリフィックスを設定 Ctrl,Shift,Alt 最大3個
                                                    size_t value = sizeof(::m_aControlKey) / sizeof(ASSOC);
                                                    for (size_t index = 0; index < value; index++)
                                                    {
                                                        LPASSOC value = &::m_aControlKey[index];
                                                        if (value->key == (source->fVirt & (value->key)))
                                                        {
                                                            ::_tcscat_s(result->dwTypeData, length, value->value);
                                                        }
                                                    };
                                                    // A～Z 文字
                                                    if (_T('A') <= source->key && source->key <= _T('Z'))
                                                    {
                                                        size_t index = 2;
                                                        LPTSTR value = (LPTSTR)::malloc(sizeof(TCHAR) * index);
                                                        if (nullptr != value)
                                                        {
                                                            ::_stprintf_s(value, index, _T("%c"), source->key);
                                                            ::_tcscat_s(result->dwTypeData, length, value);
                                                            ::free(value);
                                                            value = nullptr;
                                                        }
                                                    }
                                                    // F1～F24 文字列
                                                    else if (VK_F1 <= source->key && source->key <= VK_F24)
                                                    {
                                                        size_t index = 3;
                                                        LPTSTR value = (LPTSTR)::malloc(sizeof(TCHAR) * index);
                                                        if (nullptr != value)
                                                        {
                                                            ::SecureZeroMemory(value, sizeof(TCHAR) * index);
                                                            ::_stprintf_s(value, index, _T("F%d"), (int)(1 + source->key - VK_F1));
                                                            ::_tcscat_s(result->dwTypeData, length, value);
                                                            ::SecureZeroMemory(value, sizeof(TCHAR) * index);
                                                            ::free(value);
                                                            value = nullptr;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        // その他アクセラレーターキー
                                                        BOOL find = FALSE;
                                                        size_t value = sizeof(::m_aValueKey) / sizeof(ASSOC);
                                                        for (size_t index = 0; FALSE == find && index < value; index++)
                                                        {
                                                            LPASSOC value = &::m_aValueKey[index];
                                                            find = source->key == value->key;
                                                            if (find)
                                                            {
                                                                ::_tcscat_s(result->dwTypeData, length, value->value);
                                                            }
                                                        };
                                                    }
                                                    // メニューバーへアクセラレータ文字列を設定する
                                                    ::SetMenuItemInfo(hMenu, source->cmd, FALSE, result);
                                                }
                                                else
                                                {
                                                    // メニューアイテム文字列に既にアクセラレータ文字列が設定済みの場合
                                                    ::free((void*)value);
                                                    value = nullptr;
                                                }
                                            }// メニューアイテムに設定された文字列の取得失敗
                                        }
                                    };
                                    // メニューアイテム情報設定構造体の破棄
                                    ::free(result);
                                    result = nullptr;
                                }// メニューアイテム情報作成失敗
                            }// アクセラレータ情報の取得に失敗（読み込み個数が一致しない）
                            // アクセラレータ情報取得配列の解放
                            ::free(result);
                            result = nullptr;
                        }// アクセラレータ情報配列の作成に失敗
                    }// アクセラレータの個数が 0
                }// メニューの取得に失敗
                // メイン メッセージ ループ:
                UINT index = sizeof(MSG);
                LPMSG value = (LPMSG)::malloc(index);
                if (nullptr != value)
                {
                    ::SecureZeroMemory(value, index);
                    while (::GetMessage(value, nullptr, 0, 0))
                    {
                        if (!::TranslateAccelerator((HWND)value->hwnd, source, value))
                        {
                            ::TranslateMessage(value);
                            ::DispatchMessage(value);
                        }
                    };
                    result = (int)value->wParam;
                    // メッセージ情報構造体の破棄
                    ::free(value);
                    value = nullptr;
                }
                // アクセラレータハンドルの破棄
                ::DestroyAcceleratorTable(source);
                source = nullptr;
            }
        }
        value = (void*)::m_szTitle;
        if (nullptr != value)
        {
            // タイトル名バッファの解放
            ::free(value);
            ::m_szTitle = nullptr;
        }
        value = (void*)::m_szWindowClass;
        if (nullptr != value)
        {
            // クラス名バッファの解放
            ::free(value);
            ::m_szWindowClass = nullptr;
        }
        ::CoUninitialize();
    }
    return result;
}



//
//  関数: MyRegisterClass()
//
//  目的: ウィンドウ クラスを登録します。
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    int index = 1 + MAX_PATH;
    {
        UINT value = sizeof(TCHAR) * index;
        LPTSTR source = (LPTSTR)::malloc(value);
        if (nullptr != source)
        {
            ::SecureZeroMemory(source, value);
            if (0 <= ::LoadString(::m_hInst, IDR_MAINFRAME, source, index))
            {
                LPCTSTR* value = nullptr;
                if (3 <= ::Split(&value, (LPCTSTR)source))
                {
                    ::m_szTitle = ::_tcsdup(value[0]);
                    ::_tcscpy_s(source, index, value[2]);
                }
                ::Split(&value);
            }
            ::m_szWindowClass = (LPCTSTR)source;
            source = nullptr;
        }
    }
    ATOM result = 0;
    index = sizeof(WNDCLASSEX);
    LPWNDCLASSEX value = (LPWNDCLASSEX)::malloc(index);
    if (nullptr != value)
    {
        ::SecureZeroMemory(value, index);

        value->cbSize           = index;
        value->style            = CS_HREDRAW | CS_VREDRAW;
        value->lpfnWndProc      = ::WndProc;
        value->cbClsExtra       = 0;
        value->cbWndExtra       = 0;
        value->hInstance        = hInstance;
        value->hIcon            = ::LoadIcon(hInstance, MAKEINTRESOURCE(IDR_MAINFRAME));
        value->hCursor          = ::LoadCursor(nullptr, IDC_ARROW);
        value->hbrBackground    = (HBRUSH)(COLOR_WINDOW + 1);
        value->lpszMenuName     = MAKEINTRESOURCE(IDR_MAINFRAME);
        value->lpszClassName    = ::m_szWindowClass;
        value->hIconSm          = ::LoadIcon(hInstance, MAKEINTRESOURCE(IDR_MAINFRAME));
        result = ::RegisterClassEx(value);

        ::SecureZeroMemory(value, index);
        ::free(value);
        value = nullptr;
    }
    return result;
}

//
//   関数: InitInstance(HINSTANCE, int)
//
//   目的: インスタンス ハンドルを保存して、メイン ウィンドウを作成します
//
//   コメント:
//
//        この関数で、グローバル変数でインスタンス ハンドルを保存し、
//        メイン プログラム ウィンドウを作成および表示します。
//
BOOL InitInstance(HINSTANCE source, int index)
{
   ::m_hInst = source; // グローバル変数にインスタンス ハンドルを格納する
   // メインフレームの構築
   HWND value = ::CreateWindow(::m_szWindowClass, ::m_szTitle, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, source, nullptr);

   if (nullptr == value)
   {
      return FALSE;
   }

   // ツールバーの初期化
   ::SendMessage(value, WM_COMMAND, MAKEWPARAM(ID_VIEW_TOOLBAR, 0), 0);
   // ステータスバーの初期化
   ::SendMessage(value, WM_COMMAND, MAKEWPARAM(ID_VIEW_STATUS_BAR, 0), 0);
   // リストビューを「詳細表示」へ設定
   ::SendMessage(value, WM_COMMAND, MAKEWPARAM(ID_VIEW_DETAILS, 0), 0);
   // 表示開始時のメインフレームの大きさを設定
   ::SetWindowPos(value, nullptr, 0, 0, ::m_WindowStartUpSize.cx, ::m_WindowStartUpSize.cy, SWP_NOMOVE | SWP_NOZORDER);

   // コマンドラインに従って初期表示状態を切り替える
   ::ShowWindow(value, index);
   // メインフレームを更新する
   ::UpdateWindow(value);

   // ドキュメントなどを初期化する
   ::SendMessage(value, WM_COMMAND, MAKEWPARAM(ID_FILE_NEW, 0), 0);
   // メインフレームを保持する
   ::m_hMainWnd = value;
   // 「準備完了」メッセージの表示
   ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
   // インターバルタイマーを設定する
   ::SetTimer(value, ID_TIMER_INTERVAL, 125, nullptr);
   return TRUE;
}

//
//  関数: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  目的: メイン ウィンドウのメッセージを処理します。
//
//  WM_FINDREPLACE
//      FR_FINDNEXT
//      FR_REPLACE
//      FR_REPLACEALL
//      FR_DIALOGTERM
//  WM_CLOSE
//  WM_COMMAND  - アプリケーション メニューの処理
//      ID_FILE_NEW
//      ID_FILE_OPEN
//      ID_FILE_SAVE
//      ID_FILE_SAVE_AS
//      ID_FILE_PAGE_SETUP
//      ID_FILE_PRINT_SETUP
//      ID_FILE_PRINT_PREVIEW
//      ID_FILE_PRINT
//      ID_FILE_SEND_MAIL
//      ID_EDIT_COPY
//      ID_EDIT_CUT
//      ID_EDIT_PASTE
//      ID_EDIT_SELECT_ALL
//      ID_EDIT_UNDO
//      ID_EDIT_FIND
//      ID_EDIT_REPLACE
//      ID_EDIT_REPEAT
//      ID_OLE_INSERT_NEW
//      ID_OLE_EDIT_PROPERTIES
//      ID_EDIT_CLEAR
//      AFX_ID_PREVIEW_PRINT
//      AFX_ID_PREVIEW_CLOSE
//      AFX_ID_PREVIEW_NUMPAGE
//      AFX_ID_PREVIEW_NEXT
//      AFX_ID_PREVIEW_PREV
//      AFX_ID_PREVIEW_ZOOMIN
//      AFX_ID_PREVIEW_ZOOMOUT
//      ID_VIEW_TOOLBAR
//      ID_VIEW_STATUS_BAR
//      ID_VIEW_SMALLICON
//      ID_VIEW_LARGEICON
//      ID_VIEW_LIST
//      ID_VIEW_DETAILS
//      ID_HELP_INDEX
//      ID_HELP_FINDER
//      ID_HELP_USING
//      ID_CONTEXT_HELP
//      ID_HELP
//      ID_APP_ABOUT
//      ID_APP_EXIT
//  WM_CREATE
//  WM_DRAWITEM
//  WM_GETMINMAXINFO
//  WM_INITIALUPDATE
//  WM_KILLFOCUS
//  WM_MENUSELECT
//  WM_NOTIFY
//      TTN_NEEDTEXT
//      NM_SETFOCUS
//      NM_KILLFOCUS
//      NM_CLICK
//      LVN_ITEMCHANGED
//      LVN_KEYDOWN
//      NM_CUSTOMDRAW
//      LVN_COLUMNCLICK
//      NM_DBLCLK
//      LVN_BEGINLABELEDIT
//      LVN_ENDLABELEDIT
//  WM_PAINT    - メイン ウィンドウを描画する
//  WM_SETFOCUS
//  WM_SETMESSAGESTRING
//  WM_SIZE
//  WM_TIMER
//  WM_DESTROY  - 中止メッセージを表示して戻る
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    // 検索/置換ダイアログのコールバック処理
    if (WM_FINDREPLACE == message && nullptr != ::m_pFindReplace)
    {
        UINT value = FR_FINDNEXT;
        value |= FR_REPLACE;
        value |= FR_REPLACEALL;
        value |= FR_DIALOGTERM;
        value &= ::m_pFindReplace->Flags;
        switch (value)
        {
        case FR_FINDNEXT:
            // 「次を検索」処理
            ::SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(ID_EDIT_REPEAT, 0), 0);
            break;
        case FR_REPLACE:
        case FR_REPLACEALL:
            // 「置換して次へ/すべて置換」処理
            if (::IsWindow(hWnd))
            {
                int result = 0;
                // リストビューのハンドルを取得
                HWND source = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                if (::IsWindow(source))
                {
                    switch (value)
                    {
                    case FR_REPLACE:
                        // 「置換して次へ」処理
                        if (TRUE)
                        {
                            // 現在位置のサブアイテムの文字列を取得
                            LPCTSTR value = ::DDX_ListViewItemText(source, ::m_iSubItem);
                            if (nullptr != value)
                            {
                                // 文字列を置換する
                                LPCTSTR index = ::replace(value, ::m_pFindReplace->lpstrFindWhat, ::m_pFindReplace->lpstrReplaceWith, (LPUINT)&result);
                                if (nullptr != index)
                                {
                                    // 更新判定
                                    if (0 < result &&
                                        0 != ::_tcscmp(value, index))
                                    {
                                        // 更新文字列をサブアイテムへ書き戻す
                                        ::DDX_ListViewItemText(source, ::m_iSubItem, index);
                                        // 編集中状態に設定
                                        ::m_bDirty = TRUE;
                                    }
                                    // 更新後の文字列の解放
                                    ::free((void*)index);
                                    index = nullptr;
                                }
                                // 置換前のの文字列を解放
                                ::free((void*)value);
                                value = nullptr;
                            }
                            // 「次を検索」処理
                            ::SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(ID_EDIT_REPEAT, 0), 0);
                        }
                        break;
                    case FR_REPLACEALL:
                        // 「すべて置換」処理
                        if (TRUE)
                        {
                            int count = (int)::SendMessage(source, LVM_GETITEMCOUNT, 0u, 0u),
                                iMaxSubItems = ::GetSubItemCount(),
                                iFindItem = ::m_iItem;
                            for (int iItem = 0; iItem < count; iItem++)
                            {
                                ::m_iItem = iItem;
                                for (int iSubItem = 0; iSubItem < iMaxSubItems; iSubItem++)
                                {
                                    // 現在位置のサブアイテムの文字列を取得
                                    LPCTSTR value = ::DDX_ListViewItemText(source, iSubItem);
                                    if (nullptr != value)
                                    {
                                        int count = 0;
                                        // 文字列を置換する
                                        LPCTSTR index = ::replace(value, ::m_pFindReplace->lpstrFindWhat, ::m_pFindReplace->lpstrReplaceWith, (LPUINT)&count);
                                        if (nullptr != index)
                                        {
                                            // 更新判定
                                            if (0 < count &&
                                                0 != ::_tcscmp(value, index))
                                            {
                                                // 更新文字列をサブアイテムへ書き戻す
                                                ::DDX_ListViewItemText(source, iSubItem, index);
                                                // 置換した回数を集計加算
                                                result += (int)count;
                                                // 編集中状態に設定
                                                ::m_bDirty = TRUE;
                                            }
                                            // 更新後の文字列の解放
                                            ::free((void*)index);
                                            index = nullptr;
                                        }
                                        // 置換前のの文字列を解放
                                        ::free((void*)value);
                                        value = nullptr;
                                    }
                                };
                            };
                            ::m_iItem = iFindItem;
                        }
                        break;
                    }
                }// リストビューハンドルの取得失敗
                // 置換個数判定
                if (0 < result)
                {
                    // 「??? 個のアイテムを置換しました。」文字列を作成
                    int source = result,
                        index = 1 + MAX_PATH;
                    LPTSTR
                        result = nullptr,
                        // 個数文字列バッファを構築
                        value = (LPTSTR)::malloc(sizeof(TCHAR) * index);
                    if (nullptr != value)
                    {
                        ::SecureZeroMemory(value, sizeof(TCHAR) * index);
                        ::_stprintf_s(value, (size_t)index, _T("%d"), source);
                        result = (LPTSTR)::AfxFormatString(IDP_EDIT_REPLACE_COUNT, (LPCTSTR)value);
                        ::SecureZeroMemory(value, sizeof(TCHAR) * index);
                        // 個数文字列を解放
                        ::free(value);
                        value = nullptr;
                    }
                    if (nullptr != result)
                    {
                        // 作成した文字列を表示
                        ::MessageBox((LPCTSTR)result, MB_ICONINFORMATION);
                        // 作成した文字列を解放
                        ::free(result);
                        result = nullptr;
                    }
                }
                else
                {
                    // 「検索文字列が見つかりません。」表示
                    ::MessageBox(AFX_IDP_E_SEARCHTEXTNOTFOUND, MB_ICONINFORMATION);
                }
            }
            break;
        case FR_DIALOGTERM:
            // 「キャンセル」処理
            ::SendMessage(::m_hFindReplace, WM_DESTROY, 0u, 0u);
            ::m_hFindReplace = nullptr;
            ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
            break;
        }
    }
    switch (message)
    {
    case WM_CLOSE:
        // メインフレームの「終了」処理。もしくは、
        // 印刷プレビューの「閉じる」処理
        if (::IsWindow(hWnd))
        {
            if (::IsMenu(::m_hMenu))
            {
                // 印刷プレビューの「閉じる」処理
                // 印刷プレビューバーの非表示
                HWND value = ::GetDlgItem(hWnd, AFX_IDW_PREVIEW_BAR);
                if (::IsWindow(value))
                {
                    ::ShowWindow(value, SW_HIDE);
                }
                // ツールバーの表示
                value = ::GetDlgItem(hWnd, AFX_IDW_TOOLBAR);
                if (::IsWindow(value))
                {
                    ::ShowWindow(value, SW_SHOW);
                }
                // リストビューの表示
                value = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                if (::IsWindow(value))
                {
                    ::ShowWindow(value, SW_SHOW);
                }
                // メニューバーの復帰
                ::SetMenu(hWnd, ::m_hMenu);
                ::m_hMenu = nullptr;
            }
            else
            {
                // メインフレームの終了処理
                BOOL result = TRUE;
                // 編集中判定
                if (::m_bDirty)
                {
                    // 保存確認
                    result = ::SaveModified();
                    if (FALSE != result)
                    {
                        ::m_bDirty = FALSE;
                    }
                }
                else
                {
                    // 終了確認
                    result = IDYES == ::MessageBox(AFX_IDP_ASK_TO_EXIT, MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2);
                }
                if (result)
                {
                    // 終了判定で本当にメインフレームを終了する場合
                    // インターバルタイマーを終了する
                    ::KillTimer(hWnd, ID_TIMER_INTERVAL);
                    // データベースを閉じる
                    ::CloseDB();
                    // 印刷中メッセージ用文字列バッファの破棄
                    LPVOID value = (LPVOID)::m_pPrinterName;
                    if (nullptr != value)
                    {
                        ::free(value);
                        ::m_pPrinterName = nullptr;
                    }
                    // 起動中の「ヘルプ」画面があれば終了する
                    HWND index = ::m_hHelp;
                    if (::IsWindow(index))
                    {
                        ::SendMessage(index, WM_CLOSE, 0, 0);
                        ::m_hHelp = nullptr;
                    }
                    // 起動中の「検索/置換」ダイアログがあれば、破棄する
                    index = ::m_hFindReplace;
                    if (::IsWindow(index))
                    {
                        ::DestroyWindow(index);
                        ::m_hFindReplace = nullptr;
                    }
                    // 「検索/置換」ダイアログ用構造体を破棄
                    {
                        LPFINDREPLACE source = ::m_pFindReplace;
                        if (nullptr != source)
                        {
                            value = (LPVOID)source->lpstrReplaceWith;
                            if (nullptr != value)
                            {
                                ::free(value);
                                source->lpstrReplaceWith = nullptr;
                            }
                            value = (LPVOID)source->lpstrFindWhat;
                            if (nullptr != value)
                            {
                                ::free(value);
                                source->lpstrFindWhat = nullptr;
                            }
                            ::free(source);
                            ::m_pFindReplace = nullptr;
                        }
                    }
                    // ビュー（リストビューの破棄）
                    index = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                    if (::IsWindow(index))
                    {
                        // リストビューのサブクラスを戻す
                        WNDPROC source = ::m_wndList;
                        if (0 != source)
                        {
#ifdef _M_IX86
                            ::SetWindowLong(index, GWL_WNDPROC, (LONG)(LONGLONG)source);
#else
                            ::SetWindowLongPtr(index, GWLP_WNDPROC, (LONG_PTR)source);
#endif
                            ::m_wndList = 0;
                        }
                        HIMAGELIST result = nullptr;
                        UINT value = ::GetWindowLong(index, GWL_STYLE);
                        switch (LVS_SHAREIMAGELISTS & value)
                        {
                        case LVS_SHAREIMAGELISTS:
                            // リストビューのイメージリスト（大）を破棄
                            result = (HIMAGELIST)::SendMessage(
                                index, LVM_GETIMAGELIST, LVSIL_NORMAL, 0);
                            if (nullptr != result &&
                                INVALID_HANDLE_VALUE != result)
                            {
                                ::ImageList_Destroy(result);
                                ::SendMessage(index, LVM_SETIMAGELIST, LVSIL_NORMAL, 0);
                                result = nullptr;
                            }
                            // リストビューのイメージリスト（小）を破棄
                            result = (HIMAGELIST)::SendMessage(index, LVM_GETIMAGELIST, LVSIL_SMALL, 0);
                            if (nullptr != result &&
                                INVALID_HANDLE_VALUE != result)
                            {
                                ::ImageList_Destroy(result);
                                ::SendMessage(index, LVM_SETIMAGELIST, LVSIL_SMALL, 0);
                                result = nullptr;
                            }
                            break;
                        }
                        ::DestroyWindow(index);
                        index = nullptr;
                    }
                    // ステータスバー用許可/禁止配列の破棄
                    value = (LPVOID)::m_bStatusBar;
                    if (nullptr != value)
                    {
                        ::free(value);
                        ::m_bStatusBar = nullptr;
                    }
                    // ステータスバー用表示文字列配列の破棄
                    if (nullptr != ::m_aStatusBar)
                    {
                        for (int index = 0; nullptr != ::m_aStatusBar[index]; index++)
                        {
                            ::free((void*)::m_aStatusBar[index]);
                            ::m_aStatusBar[index] = nullptr;
                        };
                        ::free(::m_aStatusBar);
                        ::m_aStatusBar = nullptr;
                    }
                    // ステータスバーのコマンドID保持用配列の破棄
                    if (nullptr != ::m_pStatusBar)
                    {
                        ::free(::m_pStatusBar);
                        ::m_pStatusBar = nullptr;
                    }
                    // ステータスバーの破棄
                    index = ::GetDlgItem(hWnd, AFX_IDW_STATUS_BAR);
                    if (::IsWindow(index))
                    {
                        ::DestroyWindow(index);
                        index = nullptr;
                    }
                    // ツールチップ表示用文字列バッファの破棄
                    value = (LPVOID)::m_sToolTip;
                    if (nullptr != value)
                    {
                        ::free(value);
                        ::m_sToolTip = nullptr;
                    }
                    // 印刷プレビュー用ツールバーの破棄
                    index = ::GetDlgItem(hWnd, AFX_IDW_PREVIEW_BAR);
                    if (::IsWindow(index))
                    {
                        ::DestroyWindow(index);
                        index = nullptr;
                    }
                    // ツールバーの破棄
                    index = ::GetDlgItem(hWnd, AFX_IDW_TOOLBAR);
                    if (::IsWindow(index))
                    {
                        ::DestroyWindow(index);
                        index = nullptr;
                    }
                    // ページ設定ダイアログ用構造体の破棄
                    value = (LPVOID)::m_pPageSetupDlg;
                    if (nullptr != value)
                    {
                        ::free(value);
                        ::m_pPageSetupDlg = nullptr;
                    }
                    // 印刷設定ダイアログ用構造体の破棄
                    value = ::m_pPrintDlg;
                    if (nullptr != value)
                    {
                        ::free(value);
                        ::m_pPrintDlg = nullptr;
                    }
                    // ファイル名取得ダイアログ用構造体の破棄
                    {
                        LPOPENFILENAME source = ::m_pOpenFileName;
                        if (nullptr != source)
                        {
                            // ファイルフィルタ用文字列バッファの破棄
                            if (nullptr != source->lpstrFilter)
                            {
                                ::free((LPVOID)source->lpstrFilter);
                                source->lpstrFilter = nullptr;
                            }
                            // ファイルタイトル用文字列バッファの破棄
                            if (nullptr != source->lpstrFileTitle)
                            {
                                ::free(source->lpstrFileTitle);
                                source->lpstrFileTitle = nullptr;
                            }
                            // ファイル名用文字列バッファの破棄
                            if (nullptr != source->lpstrFile)
                            {
                                ::free(source->lpstrFile);
                                source->lpstrFile = nullptr;
                            }
                            ::free(source);
                            ::m_pOpenFileName = nullptr;
                        }
                    }
                    // メインフレームの破棄
                    ::DestroyWindow(hWnd);
                }
            }
        }
        break;
    case WM_COMMAND:
        if (::IsWindow(hWnd))
        {
            BOOL result = TRUE;
            UINT value = 0u;
            int index = LOWORD(wParam);
            // 選択されたメニューの解析:
            HMENU hMenu = ::GetMenu(hWnd);
            if (nullptr != hMenu && INVALID_HANDLE_VALUE != hMenu)
            {
                value = MF_GRAYED | MF_DISABLED;
                value &= ::GetMenuState(hMenu, index, MF_BYCOMMAND);
                result = MF_ENABLED == value;
            }
            if (result)
            {
                switch (index)
                {
                case ID_FILE_NEW:
                    // 「新規作成」処理
                    if (nullptr != ::m_pOpenFileName && ::SaveModified())
                    {
                        LPOPENFILENAME value = ::m_pOpenFileName;
                        if (nullptr != value->lpstrFile &&
                            NULLSTR != *value->lpstrFile)
                        {
                            *value->lpstrFile = NULLSTR;
                        }
                        if (nullptr != value->lpstrFileTitle &&
                            NULLSTR != *value->lpstrFileTitle)
                        {
                            *value->lpstrFileTitle = NULLSTR;
                        }
                        ::CloseDB();
                        ::m_bDirty = FALSE;
                        // ビューを更新する
                        ::SendMessage(hWnd, WM_INITIALUPDATE, 0, 0);
                        // 「準備完了」表示
                        ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                    }
                    break;
                case ID_FILE_OPEN:
                    // 「ドキュメントを開く」処理
                    if (nullptr != ::m_pOpenFileName &&
                        ::GetOpenFileName(::m_pOpenFileName))
                    {
                        // ドキュメントが編集中かどうかを判定する
                        if (::SaveModified())
                        {
                            // DBを開く
                            if (::OpenDB())
                            {
                                ::m_bDirty = FALSE;
                                // ビューを更新する
                                ::SendMessage(hWnd, WM_INITIALUPDATE, 0, 0);
                                // 「準備完了」表示
                                ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                            }
                            else
                            {
                                // ドキュメントを開くことに失敗しました。
                                ::MessageBox(AFX_IDP_FAILED_TO_OPEN_DOC);
                            }
                        }
                    }
                    else
                    {
                        // キャンセル
                        ::MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
                    }
                    break;
                case ID_FILE_SAVE:
                    // 「ファイルを保存」処理
                    if (::IsWindow(hWnd))
                    {
                        LPOPENFILENAME value = ::m_pOpenFileName;
                        if (nullptr != value &&
                            nullptr != value->lpstrFile &&
                            NULLSTR != *value->lpstrFile)
                        {
                            ::m_bDirty = FALSE;
                            ::MessageBox(_T("TODO: Under construction: 工事中：保存"), MB_ICONINFORMATION | MB_DEFBUTTON2);
                            ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                        }
                        else
                        {
                            // 「名前をつけて保存」へ分岐
                            ::SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(ID_FILE_SAVE_AS, 0), 0);
                        }
                    }
                    break;
                case ID_FILE_SAVE_AS:
                    // 「名前をつけて保存」処理
                    if (::IsWindow(hWnd))
                    {
                        LPOPENFILENAME value = ::m_pOpenFileName;
                        if (nullptr != value)
                        {
                            if (::GetSaveFileName(value))
                            {
                                // 「保存」処理へ分岐
                                ::SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(ID_FILE_SAVE, 0), 0);
                            }
                            else
                            {
                                // キャンセル
                                ::MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
                            }
                        }
                    }
                    break;
                case ID_FILE_PAGE_SETUP:
                    // 「ページ設定」処理
                    if (::IsWindow(hWnd))
                    {
                        LPPAGESETUPDLG value = ::m_pPageSetupDlg;
                        if (nullptr != value)
                        {
                            // ページ設定ダイアログ起動
                            value->Flags &= ~PSD_RETURNDEFAULT;
                            value->Flags |= PSD_INHUNDREDTHSOFMILLIMETERS;
                            if (::PageSetupDlg(value))
                            {
                                //入力内容を保存
                                LPCTSTR result = ::SetPrinter(value->hDevMode);
                                if (nullptr == result)
                                {
                                    // 「準備完了」表示
                                    ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                                }
                                else
                                {
                                    // エラーを表示
                                    ::MessageBox(result);
                                }
                            }
                            else
                            {
                                // キャンセル
                                ::MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
                            }
                        }
                    }
                    break;
                case ID_FILE_PRINT_SETUP:
                    // 「印刷設定」処理
                    if (::IsWindow(hWnd))
                    {
                        LPPRINTDLG value = ::m_pPrintDlg;
                        if (nullptr != value)
                        {
                            // 「印刷設定」ダイアログを起動
                            value->Flags &= ~PD_RETURNDEFAULT;
                            value->Flags |= PD_PRINTSETUP;
                            value->lpfnPrintHook = nullptr;
                            if (::PrintDlg(value))
                            {
                                // 入力内容を保存
                                LPCTSTR result = ::SetPrinter(value->hDevMode);
                                if (nullptr == result)
                                {
                                    // 「準備完了」を表示
                                    ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                                }
                                else
                                {
                                    // エラーメッセージを表示
                                    ::MessageBox(result);
                                }
                            }
                            else
                            {
                                // キャンセル
                                ::MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
                            }
                        }
                    }
                    break;
                case ID_FILE_PRINT_PREVIEW:
                    // 「印刷プレビュー」処理
                    if (::IsWindow(hWnd))
                    {
                        // ツールバーを非表示
                        HWND value = ::GetDlgItem(hWnd, AFX_IDW_TOOLBAR);
                        if (::IsWindow(value))
                        {
                            ::ShowWindow(value, SW_HIDE);
                        }
                        // ステータスバーを非表示
                        value = ::GetDlgItem(hWnd, AFX_IDW_PREVIEW_BAR);
                        if (::IsWindow(value))
                        {
                            ::ShowWindow(value, SW_SHOW);
                        }
                        // リストビューを非表示
                        value = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                        if (::IsWindow(value))
                        {
                            ::ShowWindow(value, SW_HIDE);
                        }
                        // メニューバーを保持して破棄する
                        ::m_hMenu = ::GetMenu(hWnd);
                        ::SetMenu(hWnd, nullptr);

                        // 印刷プレビューの設定
                        ::m_nPrintMarginLeft    = 50;   // 仮
                        ::m_nPrintMarginTop     = 38;   // 仮
                        ::m_nPrintLineByPage    = 40;   // 仮
                        ::m_nPrintMaxPage       = 3;    // 仮
                        ::m_nPrintRowPitch      = 19;   // 仮
                        ::m_nPrintFieldPitch    = 100;  // 仮
                        ::m_nPrintCurrentPage   = 1;
                        ::m_nPreviewpageNum     = 1;
                        ::InvalidateRect(hWnd, nullptr, TRUE);
                    }
                    break;
                case ID_FILE_PRINT:
                    // 「印刷」処理
                    if (::IsWindow(hWnd) && nullptr != ::m_pPrintDlg)
                    {
                        LPPRINTDLG value        = ::m_pPrintDlg;
                        ::m_nPrintMarginLeft    = 100;  // 仮
                        ::m_nPrintMarginTop     = 100;  // 仮
                        ::m_nPrintLineByPage    = 40;   // 仮
                        ::m_nPrintMaxPage       = 3;    // 仮
                        ::m_nPrintRowPitch      = 100;  // 仮
                        ::m_nPrintFieldPitch    = 600;  // 仮
                        // 「印刷」ダイアログの表示
                        value->Flags
                            &= ~(PD_PRINTSETUP | PD_RETURNDEFAULT);
                        value->Flags
                            |= PD_PAGENUMS
                            | PD_RETURNDC
                            | PD_USEDEVMODECOPIES
                            | PD_COLLATE
                            | PD_NOSELECTION
                            | PD_ENABLEPRINTHOOK
                            ;
                        value->nCopies = 1;
                        value->nMinPage = 1;
                        value->nMaxPage = ::m_nPrintMaxPage;
                        value->nFromPage = value->nMinPage;
                        value->nToPage = value->nMaxPage;
                        value->lpfnPrintHook = ::PrintSetupProc;
                        switch (::PrintDlg(value))
                        {
                        case IDOK:
                            // 「印刷中」ダイアログの表示
                            switch (::DialogBox(::m_hInst, MAKEINTRESOURCE(AFX_IDD_PRINTDLG), hWnd, ::PrintProc))
                            {
                            case IDOK:
                                // 「準備完了」表示
                                ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                                break;
                            default:
                                // 印刷中をキャンセル
                                ::MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
                                break;
                            }
                            break;
                        default:
                            // 印刷ダイアログをキャンセル
                            ::MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
                            break;
                        }
                    }
                    break;
                case ID_FILE_SEND_MAIL:
                    // 「送信」処理
                    ::MessageBox(_T("TODO: Under construction: 工事中：送信"), MB_ICONINFORMATION | MB_DEFBUTTON2);
                    ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                    break;
                case ID_EDIT_COPY:
                case ID_EDIT_CUT:
                case ID_EDIT_PASTE:
                case ID_EDIT_SELECT_ALL:
                case ID_EDIT_UNDO:
                    if (::IsWindow(hWnd))
                    {
                        // リストビューハンドルの取得
                        HWND source = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                        if (::IsWindow(source))
                        {
                            // フローティングエディタのハンドルの取得
                            source = (HWND)::SendMessage(source, LVM_GETEDITCONTROL, 0, 0);
                            if (::IsWindow(source))
                            {
                                UINT result = 0;
                                value = 0;
                                switch (index)
                                {
                                case ID_EDIT_COPY:
                                    // 「コピー」処理
                                    result = WM_COPY;
                                    break;
                                case ID_EDIT_CUT:
                                    // 「切り取り」処理
                                    result = WM_CUT;
                                    break;
                                case ID_EDIT_PASTE:
                                    // 「貼り付け」処理
                                    result = WM_PASTE;
                                    break;
                                case ID_EDIT_SELECT_ALL:
                                    // 「すべて選択」処理
                                    value = ::GetWindowTextLength(source);
                                    result = EM_SETSEL;
                                    break;
                                case ID_EDIT_UNDO:
                                    // 「元に戻す」処理
                                    result = EM_UNDO;
                                    ::SendMessage(source, EM_UNDO, 0, 0);
                                    break;
                                }
                                if (0 != result)
                                {
                                    ::SendMessage(source, result, 0, (LPARAM)value);
                                }
                                // 「準備完了」表示
                                ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                            }// フローティングエディタが起動していない
                        }// リストビューハンドルの取得失敗
                    }// フレームウィンドウが無効
                    break;
                case ID_EDIT_FIND:
                case ID_EDIT_REPLACE:
                    // 「検索/置換」ダイアログ起動処理
                    if (FALSE == ::IsWindow(::m_hFindReplace))
                    {
                        LPFINDREPLACE source = ::m_pFindReplace;
                        if (nullptr != source)
                        {
                            HWND result = nullptr;
                            source->Flags = FR_DOWN;
                            switch (index)
                            {
                            case ID_EDIT_FIND:
                                // 「検索」ダイアログ起動処理
                                result = ::FindText(source);
                                value = IDP_EDIT_FIND_STRING;
                                break;
                            case ID_EDIT_REPLACE:
                                // 「置換」ダイアログ起動処理
                                result = ::ReplaceText(source);
                                value = IDP_EDIT_REPLACE_STRING;
                                break;
                            }
                            if (FALSE != ::IsWindow(result))
                            {
                                ::m_hFindReplace = result;
                                // 該当指示文字列の表示
                                ::MessageBox(value, MB_ICONINFORMATION);
                            }
                            break;
                        }
                    }
                    break;
                case ID_EDIT_REPEAT:
                    // 「次を検索」処理
                    if (nullptr != ::m_pFindReplace)
                    {
                        HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                        if (::IsWindow(hCtrl))
                        {
                            int iFindItem = ::m_iItem;
                            int iMaxItems = (int)::SendMessage(hCtrl, LVM_GETITEMCOUNT, 0u, 0u);
                            int iMaxSubItems = ::GetSubItemCount();
                            result = FALSE;
                            switch (::m_pFindReplace->Flags & FR_DOWN)
                            {
                            case FR_DOWN:
                                // 現在位置から最後に向けて検索
                                for (int iItem = iFindItem;
                                    FALSE == result && iItem < iMaxItems;
                                    iItem++)
                                {
                                    ::m_iItem = iItem;
                                    int value = iFindItem == iItem ? ::m_iSubItem : 0;
                                    for (int iSubItem = value; FALSE == result && iSubItem < iMaxSubItems; iSubItem++)
                                    {
                                        result = ::FindItem(hCtrl, iSubItem);
                                        if (FALSE != result)
                                        {
                                            ::m_iSubItem = iSubItem;
                                        }
                                    };
                                };
                                if (FALSE == result)
                                {
                                    // 先頭から現在位置に向けて検索
                                    for (int iItem = 0; FALSE == result && iItem <= iFindItem; iItem++)
                                    {
                                        ::m_iItem = iItem;
                                        int value = iFindItem == iItem ? ::m_iSubItem : iMaxSubItems;
                                        for (int iSubItem = 0; FALSE == result && iSubItem < value; iSubItem++)
                                        {
                                            result = ::FindItem(hCtrl, iSubItem);
                                            if (FALSE != result)
                                            {
                                                ::m_iSubItem = iSubItem;
                                            }
                                        };
                                    };
                                }
                                break;
                            default:
                                // 現在位置から先頭に向けて検索
                                for (int iItem = iFindItem; FALSE == result && 0 <= iItem; iItem--)
                                {
                                    ::m_iItem = iItem;
                                    int value = iFindItem == iItem ? ::m_iSubItem : iMaxSubItems - 1;
                                    for (int iSubItem = value; FALSE == result && 0 <= iSubItem; iSubItem--)
                                    {
                                        result = ::FindItem(hCtrl, iSubItem);
                                        if (FALSE != result)
                                        {
                                            ::m_iSubItem = iSubItem;
                                        }
                                    };
                                };
                                if (FALSE == result)
                                {
                                    // 末尾から現在位置に向けて検索
                                    for (int iItem = iMaxItems - 1; FALSE == result && iFindItem <= iItem; iItem--)
                                    {
                                        ::m_iItem = iItem;
                                        int value = iFindItem == iItem ? ::m_iSubItem : 0;
                                        for (int iSubItem = iMaxSubItems - 1; FALSE == result && value <= iSubItem; iSubItem--)
                                        {
                                            result = ::FindItem(hCtrl, iSubItem);
                                            if (FALSE != result)
                                            {
                                                ::m_iSubItem = iSubItem;
                                            }
                                        };
                                    };
                                }
                                break;
                            }
                            if (FALSE != result)
                            {
                                // 発見されたs場合
                                ::SetRowCol(hCtrl);
                                ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                                ::SetTimer(hWnd, ID_TIMER_REEDIT, 125, nullptr);
                            }
                            else
                            {
                                // 見つからなかった場合
                                ::m_iItem = iFindItem;
                                ::MessageBox(AFX_IDP_E_SEARCHTEXTNOTFOUND, MB_ICONINFORMATION);
                            }
                        }
                    }
                    break;
                case ID_OLE_INSERT_NEW:
                    // 「挿入」処理
                    ::MessageBox(_T("TODO: Under construction: 工事中：挿入"), MB_ICONINFORMATION | MB_DEFBUTTON2);
                    break;
                case ID_OLE_EDIT_PROPERTIES:
                    // 「編集」処理
                    if (::IsWindow(hWnd))
                    {
                        HWND value = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                        if (::IsWindow(value))
                        {
                            ::SetRowCol(value);
                            ::SendMessage(value, LVM_EDITLABEL, (WPARAM)::m_iItem, 0);
                        }
                    }
                    break;
                case ID_EDIT_CLEAR:
                    // 「削除」処理
                    ::MessageBox(_T("TODO: Under construction: 工事中：削除"), MB_ICONINFORMATION);
                    break;
                case AFX_ID_PREVIEW_PRINT:
                case AFX_ID_PREVIEW_CLOSE:
                    // 印刷プレビュー中の「印刷/終了」処理
                    ::SendMessage(hWnd, WM_CLOSE, 0, 0);
                    switch (index)
                    {
                    case AFX_ID_PREVIEW_PRINT:
                        // 印刷プレビュー中の「終了」処理
                        ::SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(ID_FILE_PRINT, 0), 0);
                        break;
                    }
                    ::m_iItem = ::m_iSubItem = 0;
                    break;
                case AFX_ID_PREVIEW_NUMPAGE:
                    // 印刷プレビュー中の「1ページ/2ページ切り替え」処理
                    if (::IsWindow(hWnd))
                    {
                        UINT source = AFX_IDS_ONEPAGE;
                        value = ::m_nPreviewpageNum;
                        switch (value)
                        {
                        case 1:
                            // 「1ページ→2ページ切り替え」処理
                            value = 2;
                            source = AFX_IDS_TWOPAGE;
                            break;
                        default:
                            // 「2ページ→1ページ切り替え」処理
                            value = 1;
                            break;
                        }
                        ::m_nPreviewpageNum = value;
                        value = 1 + MAX_PATH;
                        if (0 < ::LoadString(::m_hInst, source, (LPTSTR)::m_szWindowClass, value))
                        {
                            HWND source = ::GetDlgItem(hWnd, AFX_IDW_PREVIEW_BAR);
                            if (::IsWindow(source))
                            {
                                value = sizeof(TBBUTTONINFO);
                                LPTBBUTTONINFO result = (LPTBBUTTONINFO)::malloc(value);
                                if (nullptr != result)
                                {
                                    ::SecureZeroMemory(result, value);
                                    result->cbSize = value;
                                    result->dwMask = TBIF_TEXT;
                                    result->pszText = (LPTSTR)::m_szWindowClass;
                                    if (FALSE != ::SendMessage(source, TB_SETBUTTONINFO, index, (LPARAM)result))
                                    {
                                        ::InvalidateRect(hWnd, nullptr, TRUE);
                                    }
                                    ::SecureZeroMemory(result, value);
                                    ::free(result);
                                    result = nullptr;
                                }
                            }
                        }
                    }
                    break;
                case AFX_ID_PREVIEW_NEXT:
                    // 印刷プレビュー中の「次へ」処理
                    if (::m_nPrintCurrentPage < ::m_nPrintMaxPage)
                    {
                        ::m_nPrintCurrentPage++;
                        ::InvalidateRect(hWnd, nullptr, TRUE);
                    }
                    break;
                case AFX_ID_PREVIEW_PREV:
                    // 印刷プレビュー中の「前へ」処理
                    if (1 < ::m_nPrintCurrentPage)
                    {
                        ::m_nPrintCurrentPage--;
                        ::InvalidateRect(hWnd, nullptr, TRUE);
                    }
                    break;
                case AFX_ID_PREVIEW_ZOOMIN:
                    // 印刷プレビュー中の「拡大」処理
                    ::MessageBox(_T("TODO:工事中: 印刷プレビュー拡大"), MB_ICONINFORMATION | MB_DEFBUTTON2);
                    break;
                case AFX_ID_PREVIEW_ZOOMOUT:
                    // 印刷プレビュー中の「縮小」処理
                    ::MessageBox(_T("TODO:工事中: 印刷プレビュー縮小"), MB_ICONINFORMATION | MB_DEFBUTTON2);
                    break;
                case ID_VIEW_TOOLBAR:
                case ID_VIEW_STATUS_BAR:
                    // 「ツールバー/ステータスバー」処理
                    if (::IsWindow(hWnd))
                    {
                        switch (index)
                        {
                        case ID_VIEW_TOOLBAR:
                            // 「ツールバー」処理
                            value = AFX_IDW_TOOLBAR;
                            break;
                        case ID_VIEW_STATUS_BAR:
                            // 「ステータスバー」処理
                            value = AFX_IDW_STATUS_BAR;
                            break;
                        }
                        HWND source = ::GetDlgItem(hWnd, value);
                        if (::IsWindow(source))
                        {
                            // 現在のコントロールの表示状態から許可禁止を決定
                            result = TRUE;
                            if (::IsWindow(source))
                            {
                                result = FALSE == ::IsWindowVisible(source);
                            }
                            ::ShowWindow(source, result ? SW_SHOW : SW_HIDE);
                            // メニューアイテムにチェックをつける
                            ::CheckMenuItem(hMenu, index, result ? MF_CHECKED : MF_UNCHECKED);
                            // コントロールを再配置する
                            RECT rc = { 0,0,0,0 };
                            ::GetClientRect(hWnd, &rc);
                            ::SendMessage(hWnd, WM_SIZE, 0, MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top));
                            // 「準備完了」表示
                            ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                        }
                    }
                    break;
                case ID_VIEW_SMALLICON:
                case ID_VIEW_LARGEICON:
                case ID_VIEW_LIST:
                case ID_VIEW_DETAILS:
                    // 「大きいアイコン/小さいアイコン/一覧/詳細」処理
                    value = LVS_REPORT;
                    switch (index)
                    {
                    case ID_VIEW_SMALLICON:
                        // 「小さいアイコン」処理
                        value = LVS_SMALLICON;
                        break;
                    case ID_VIEW_LARGEICON:
                        // 「大きいアイコン」処理
                        value = LVS_ICON;
                        break;
                    case ID_VIEW_LIST:
                        // 「一覧」処理
                        value = LVS_LIST;
                        break;
                    }
                    // ビューの表示切り替え処理
                    if (::IsWindow(hWnd))
                    {
                        HWND source = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                        if (::IsWindow(source))
                        {
                            ::SendMessage(source, LVM_SETVIEW, (WPARAM)value, 0);
                        }
                    }
                    {
                        value = sizeof(TBBUTTONINFO);
                        LPTBBUTTONINFO source = (LPTBBUTTONINFO)::malloc(value);
                        if (nullptr != source)
                        {
                            ::SecureZeroMemory(source, value);
                            source->cbSize = value;
                            source->dwMask = TBIF_STATE;

                            // ツールバーを取得
                            HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDW_TOOLBAR);
                            // ループする回数を算定」
                            size_t result = sizeof(::m_cuListViewCommands) / sizeof(UINT);
                            for (UINT value = 0; value < result; value++)
                            {
                                UINT result = 0;
                                {
                                    UINT source =
                                        ::m_cuListViewCommands[value];
                                    {
                                        // メニューバーの表示状態を決定
                                        result = MF_UNCHECKED;
                                        if (source == index)
                                        {
                                            result = MF_CHECKED;
                                        }
                                        // メニューバーの該当アイテムに
                                        // チェックをつける
                                        ::CheckMenuItem(hMenu, source, result);
                                        // ツールバーの表示状態を決定
                                        result = TBSTATE_ENABLED;;
                                        if (source == index)
                                        {
                                            result |= TBSTATE_CHECKED;
                                        }
                                    }
                                }
                                // ツールバーの該当ボタンにチェックをつける
                                if (::IsWindow(hCtrl))
                                {
                                    source->fsState = result;
                                    ::SendMessage(hCtrl, TB_SETBUTTONINFO, ::m_cuListViewCommands[value], (LPARAM)source);
                                }
                            };

                            ::SecureZeroMemory(source, value);
                            ::free(source);
                            source = nullptr;
                            // 「準備完了」表示
                            ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                        }
                    }
                    break;
                case ID_HELP_INDEX:
                case ID_HELP_FINDER:
                case ID_HELP_USING:
                case ID_CONTEXT_HELP:
                case ID_HELP:
                    // 「ヘルプ」処理
                    if (::IsWindow(hWnd))
                    {
                        value = 0;
                        switch (index)
                        {
                        case ID_HELP_INDEX:
                            // 「索引」処理
                            value = HH_DISPLAY_INDEX;
                            break;
                        case ID_HELP_FINDER:
                            // 「ヘルプ検索」処理
                            value = HH_DISPLAY_SEARCH;
                            break;
                        case ID_HELP_USING:
                            // 「ヘルプの使い方」処理
                            value = HH_DISPLAY_TOC;
                            break;
                        case ID_CONTEXT_HELP:
                            // 「目次」処理
                            value = HH_DISPLAY_TOPIC;
                            break;
                        case ID_HELP:
                            // 「ヘルプ」処理
                            value = HH_DISPLAY_TOC;
                            break;
                        }
                        if (0 != value)
                        {
                            if (0 < ::LoadString(::m_hInst, IDS_HELP_FILE_NAME,
                                (LPTSTR)::m_szWindowClass, 1 + MAX_PATH))
                            {
                                ::m_hHelp = ::HtmlHelp(hWnd, ::m_szWindowClass,
                                    value, 0);
                                if (::IsWindow(::m_hHelp))
                                {
                                    // 「準備完了」表示
                                    ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                                }
                            }
                        }
                    }
                    break;
                case ID_APP_ABOUT:
                    // 「このアプリケーションについて」処理
                    ::DialogBox(::m_hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, ::About);
                    ::MessageBox(AFX_IDS_IDLEMESSAGE, MB_ICONINFORMATION);
                    break;
                case ID_APP_EXIT:
                    // 「アプリケーションの終了」処理
                    ::SendMessage(hWnd, WM_CLOSE, 0u, 0u);
                    break;
                default:
                    return ::DefWindowProc(hWnd, message, wParam, lParam);
                }
            }// メニューバーアイテムが禁止状態
        }
        else
        {
            return ::DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    case WM_CREATE:
        if (::IsWindow(hWnd))
        {
            BOOL
                result = FALSE;
            UINT
                index = 0,
                value = sizeof(::m_cuSystemIcons) / sizeof(UINT);
            // イメージリストの構築処理。
            // メッセージボックスとリストビューで共用する
            HIMAGELIST
                hLargeIcon = nullptr,
                hSmallIcon = ::ImageList_Create(
                    16, 15, ILC_COLOR32 | ILC_MASK, value, 0);
            if (nullptr != hSmallIcon && INVALID_HANDLE_VALUE != hSmallIcon)
            {
                hLargeIcon = ::ImageList_Create(32, 32, ILC_COLOR32 | ILC_MASK, value, 0);
                if (nullptr != hLargeIcon &&
                    INVALID_HANDLE_VALUE != hLargeIcon)
                {
                    result = TRUE;
                    for (index = 0; result && index < value; index++)
                    {
                        result = FALSE;
                        HICON hIcon = ::LoadIcon(nullptr, MAKEINTRESOURCE(::m_cuSystemIcons[index]));
                        if (nullptr != hIcon && INVALID_HANDLE_VALUE != hIcon)
                        {
                            if (0 <= ::ImageList_AddIcon(hSmallIcon, hIcon))
                            {
                                result = 0 <= ::ImageList_AddIcon(hLargeIcon, hIcon);
                            }
                        }
                    };
                }
            }
            if (result)
            {
                result = FALSE;
                // ファイル名入出力ダイアログ用構造体作成
                value = sizeof(OPENFILENAME);
                LPOPENFILENAME source = (LPOPENFILENAME)::malloc(value);
                if (nullptr != source)
                {
                    ::SecureZeroMemory(source, value);
                    source->lStructSize = value;
                    source->hwndOwner   = hWnd;

                    // ファイル名文字列用バッファ作成
                    index = 1 + MAX_PATH;
                    value = sizeof(TCHAR) * index;
                    source->nMaxFile = index;
                    LPTSTR result = (LPTSTR)::malloc(value);
                    if (nullptr != result)
                    {
                        ::SecureZeroMemory(result, value);
                        if (0 < ::LoadString(::m_hInst, IDR_MAINFRAME,
                            result, index))
                        {
                            LPCTSTR* value = nullptr;
                            if (4 < ::Split(&value, result))
                            {
                                ::_tcscpy_s(result, index, value[3]);
                                ::_tcscat_s(result, index, _T("|*"));
                                ::_tcscat_s(result, index, value[4]);
                                ::_tcscat_s(result, index, _T("|"));
                            }
                            ::Split(&value);
                        }
                        source->lpstrFile = result;
                    }
                    // ファイルタイトル文字列用バッファ作成
                    source->nMaxFileTitle = index;
                    result = (LPTSTR)::malloc(value);
                    if (nullptr != result)
                    {
                        ::SecureZeroMemory(result, value);
                        if (0 < ::LoadString(::m_hInst, AFX_IDS_ALLFILTER,
                            result, index))
                        {
                            // ファイルフィルタを構築
                            ::_tcscat_s(source->lpstrFile, index, result);
                            ::_tcscat_s(source->lpstrFile, index, _T("|*.*|"));
                            source->lpstrFilter = ::_tcsdup(source->lpstrFile);
                            source->nFilterIndex = 1;
                            source->lpstrFileTitle = result;
                        }
                        // ファイルタイトル用バッファを空にする
                        ::SecureZeroMemory(result, value);
                    }
                    // ファイルフィルタが構築されていれば、書式を変更する
                    for (LPCTSTR index = source->lpstrFilter;
                        nullptr != index && NULLSTR != *index;
                        index = ::_tcsinc(index))
                    {
                        if (_T('|') == *index)
                        {
                            *((LPTSTR)index) = NULLSTR;
                        }
                    };
                    // ファイル名用バッファを空にする
                    ::SecureZeroMemory(source->lpstrFile, value);
                    // 構築した構造体を保持する
                    ::m_pOpenFileName = source;
                }
                result = nullptr != ::m_pOpenFileName;
            }
            if (result)
            {
                result = FALSE;
                // 印刷ダイアログ用構造体の構築
                value = sizeof(PRINTDLG);
                LPPRINTDLG source = (LPPRINTDLG)::malloc(value);
                if (nullptr != source)
                {
                    ::SecureZeroMemory(source, value);
                    source->lStructSize = value;
                    source->hwndOwner   = hWnd;
                    source->Flags       |= PD_RETURNDEFAULT;
                    result = ::PrintDlg(source);
                    ::m_pPrintDlg = source;
                }
            }
            if (result)
            {
                result = FALSE;
                // ページ設定ダイアログ用構造体の構築
                value = sizeof(PAGESETUPDLG);
                LPPAGESETUPDLG source = (LPPAGESETUPDLG)::malloc(value);
                if (nullptr != source)
                {
                    ::SecureZeroMemory(source, value);
                    source->lStructSize = value;
                    source->hwndOwner   = hWnd;
                    source->Flags       |= PSD_RETURNDEFAULT;
                    result = ::PageSetupDlg(source);
                    ::m_pPageSetupDlg = source;
                }
            }
            if (result)
            {
                result = FALSE;
                // ツールバーの構築
                value = WS_CHILD | WS_VISIBLE;
                value |= TBSTYLE_TOOLTIPS;
                value |= TBSTYLE_WRAPABLE;
                HWND hCtrl = ::CreateWindow(TOOLBARCLASSNAME, nullptr, value, 0, 0, 0, 0, hWnd, (HMENU)AFX_IDW_TOOLBAR, ::m_hInst, nullptr);
                if (::IsWindow(hCtrl))
                {
                    ::SendMessage(hCtrl, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);

                    SIZE size = { 0,0 };
                    UINT count = 0, icon = 0;
                    LPTBBUTTON pTbButtons = nullptr;
                    // ツールバーリソースを取得する
                    HRSRC hRsrc = ::FindResource(::m_hInst, MAKEINTRESOURCE(IDR_MAINFRAME), (LPTSTR)RT_TOOLBAR);
                    if (hRsrc != nullptr)
                    {
                        // 見つかったリソースハンドルをロードしてメモリーハンドルとする
                        HGLOBAL hGlobal = ::LoadResource(::m_hInst, hRsrc);
                        if (hGlobal != nullptr)
                        {
                            // メモリーハンドルをロックしてローカルメモリーを取得する
                            LPTOOLBARDATA pData = (LPTOOLBARDATA)::LockResource(hGlobal);
                            if (pData != nullptr && pData->version == 1)
                            {
                                size.cx = pData->width;
                                size.cy = pData->height;
                                count = pData->count;
                                WORD* pItems = (WORD*)&pData[1];
                                value = sizeof(TBBUTTON) * count;
                                pTbButtons = (LPTBBUTTON)::malloc(value);
                                if (nullptr != pTbButtons)
                                {
                                    ::SecureZeroMemory(pTbButtons, value);
                                    result = TRUE;
                                    for (index = 0;
                                        result && index < count;
                                        index++)
                                    {
                                        TBBUTTON& rec = pTbButtons[index];
                                        rec.fsState = TBSTATE_ENABLED;
                                        rec.fsStyle = TBSTYLE_SEP;
                                        rec.idCommand = pItems[index];
                                        switch (rec.idCommand)
                                        {
                                        case ID_SEPARATOR:
                                            break;
                                        default:
                                            rec.fsStyle = TBSTYLE_BUTTON;
                                            rec.iBitmap = icon++;
                                            if (0 < ::LoadString(::m_hInst, rec.idCommand, (LPTSTR)::m_szWindowClass, 1 + MAX_PATH))
                                            {
                                                LPCTSTR* aArray = nullptr;
                                                ::Split(&aArray, ::m_szWindowClass);
                                                rec.iString = (INT_PTR)::SendMessage(hCtrl, TB_ADDSTRING, 0, (LPARAM)aArray[2]);
                                                ::Split(&aArray);
                                            }
                                            break;
                                        }
                                    };
                                }
                                // ロックしたグローバルメモリーハンドルをアンロックする
                                UnlockResource(hGlobal);
                                hGlobal = nullptr;
                            }
                        }
                    }
                    if (nullptr != pTbButtons)
                    {
                        // リソース上のツールバービットマップをツールバーへコピーする
                        value = sizeof(TBADDBITMAP);
                        LPTBADDBITMAP pTbAddBitmap = (LPTBADDBITMAP)::malloc(value);
                        if (nullptr != pTbAddBitmap)
                        {
                            pTbAddBitmap->hInst = ::m_hInst;
                            pTbAddBitmap->nID = IDR_MAINFRAME;
                            ::SendMessage(hCtrl, TB_ADDBITMAP, (WPARAM)icon, (LPARAM)pTbAddBitmap);
                            ::free(pTbAddBitmap);
                            pTbAddBitmap = nullptr;
                        }
                        ::SendMessage(hCtrl, TB_ADDBUTTONS, (WPARAM)count, (LPARAM)pTbButtons);
                        ::SendMessage(hCtrl, TB_AUTOSIZE, 0, 0);
                        // ツールバーのアイコン画像をビットマップ画像としてメニューバーへコピーする
                        HMENU hMenu = ::GetMenu(hWnd);
                        if (::IsMenu(hMenu))
                        {
                            // メニューアイテム設定用構造体を構築する
                            value = sizeof(MENUITEMINFO);
                            LPMENUITEMINFO pMenuItemInfo = (LPMENUITEMINFO)::malloc(value);
                            if (nullptr != pMenuItemInfo)
                            {
                                ::SecureZeroMemory(pMenuItemInfo, value);
                                pMenuItemInfo->cbSize = value;
                                pMenuItemInfo->fMask = MIIM_BITMAP;

                                // イメージリスト描画用構造体を作成する
                                value = sizeof(IMAGELISTDRAWPARAMS);
                                LPIMAGELISTDRAWPARAMS source = (LPIMAGELISTDRAWPARAMS)::malloc(value);
                                if (nullptr != source)
                                {
                                    ::SecureZeroMemory(source, value);
                                    source->cbSize = value;
                                    // ツールバーのイメージリストを設定する
                                    source->himl = (HIMAGELIST)::SendMessage(hCtrl, TB_GETIMAGELIST, 0, 0);
                                    if (nullptr != source->himl &&
                                        INVALID_HANDLE_VALUE != source->himl)
                                    {
                                        // メインフォームのデバイスコンテキストを取得する
                                        HDC hDC = ::GetDC(hWnd);
                                        if (nullptr != hDC &&
                                            INVALID_HANDLE_VALUE != hDC)
                                        {
                                            source->hdcDst = ::CreateCompatibleDC(hDC);
                                            if (nullptr != source->hdcDst &&
                                                INVALID_HANDLE_VALUE != source->hdcDst)
                                            {
                                                // ツールバーび設定されているビットマップの大きさを設定
                                                source->cx      = size.cx;
                                                source->cy      = size.cy;
                                                // 描画背景色をせっいぇい
                                                source->rgbBk   = ::GetSysColor(COLOR_MENU);
                                                source->rgbFg   = CLR_NONE;
                                                source->dwRop   = SRCCOPY;

                                                // ツールバーリソースに設定されているボタンんとセパレータの数分、ループする
                                                result = TRUE;
                                                for (UINT index = 0; result && index < count; index++)
                                                {
                                                    // ルールバーのボタン情報を取得する
                                                    TBBUTTON& tbButton = pTbButtons[index];
                                                    switch (tbButton.idCommand)
                                                    {
                                                    case ID_SEPARATOR:
                                                        // セパレータをスキップする（続行：result = TRUE）
                                                        break;
                                                    default:
                                                        result = FALSE;
                                                        // メインフォーム互換のビットマップを作成する
                                                        pMenuItemInfo->hbmpItem = ::CreateCompatibleBitmap(hDC, size.cx, size.cy);
                                                        if (nullptr != pMenuItemInfo->hbmpItem &&
                                                            INVALID_HANDLE_VALUE != pMenuItemInfo->hbmpItem)
                                                        {
                                                            // 描画デバイスコンテキストへビットマップを設定する
                                                            HBITMAP hOld = (HBITMAP)::SelectObject(source->hdcDst, pMenuItemInfo->hbmpItem);
                                                            // イメージリストの描画位置を設定
                                                            source->i = tbButton.iBitmap;
                                                            // ビットマップへツールバーアイコンを描画
                                                            result = ::ImageList_DrawIndirect(source);
                                                            // 描画デバイスコンテキストからビットマップをメニューバー設定構造体へ設定する
                                                            pMenuItemInfo->hbmpItem = (HBITMAP)::SelectObject(source->hdcDst, hOld);
                                                        }
                                                        if (result)
                                                        {
                                                            // メニューバーの指定アイテムへビットマップを設定する
                                                            result = ::SetMenuItemInfo(hMenu, tbButton.idCommand, FALSE, pMenuItemInfo);
                                                        }
                                                        break;
                                                    }
                                                }
                                                // 描画用デバイスコンテキストを削除する
                                                ::DeleteDC(source->hdcDst);
                                            }
                                            // ビットマップ作成用のメインフォームのデバイスコンテキストを解放する
                                            ::ReleaseDC(hWnd, hDC);
                                        }
                                    }
                                    // イメージリスト描画用構造体を削除する
                                    ::free(source);
                                    source = nullptr;
                                }
                                // メニューアイテム設定用構造体を削除する
                                ::free(pMenuItemInfo);
                                pMenuItemInfo = nullptr;
                            }
                        }
                        // ツールバー用ボタン構造体を削除する
                        ::free(pTbButtons);
                        pTbButtons = nullptr;
                    }
                    if (result)
                    {
                        // ツールバー用ツールチップ表示用文字列バッファを構築しておく
                        ::m_sToolTip = (LPTSTR)::malloc(sizeof(TCHAR) * (size_t)(1 + MAX_PATH));
                    }
                }
            }
            if (result)
            {
                result = FALSE;
                // 印刷プレビュー用ツールバーの構築
                value = WS_CHILD /* | WS_VISIBLE*/;
                value |= TBSTYLE_LIST;
                value |= TBSTYLE_TOOLTIPS;
                value |= TBSTYLE_WRAPABLE;
                HWND hCtrl = ::CreateWindow(TOOLBARCLASSNAME, nullptr, value, 0, 0, 0, 0, hWnd, (HMENU)AFX_IDW_PREVIEW_BAR, ::m_hInst, nullptr);
                if (::IsWindow(hCtrl))
                {
                    ::SendMessage(hCtrl, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
                    // 印刷プレビュー用ツールバーのセパレータとボタンの個数を取得する
                    UINT count = sizeof(::m_cuPreviewBar) / sizeof(UINT);
                    // ツールバーボタン構造体を構築する
                    value = sizeof(TBBUTTON) * count;
                    LPTBBUTTON pTbButtons = (LPTBBUTTON)::malloc(value);
                    if (nullptr != pTbButtons)
                    {
                        ::SecureZeroMemory(pTbButtons, value);

                        // ビットマップのオフセット位置を作成する
                        UINT offset[] = { 0,0 };
                        UINT nCount = sizeof(::m_cuPreviewBarIcons) / sizeof(UINT);
                        TBADDBITMAP TbAddBitmap = { HINST_COMMCTRL,IDB_STD_SMALL_COLOR };
                        offset[0] = (UINT)::SendMessage(hCtrl, TB_ADDBITMAP, 0, (LPARAM)&TbAddBitmap);
                        TbAddBitmap.nID = IDB_HIST_SMALL_COLOR;
                        offset[1] = (UINT)::SendMessage(hCtrl, TB_ADDBITMAP, 0, (LPARAM)&TbAddBitmap);

                        // ボタンとセパレータの数分だけループする
                        for (UINT index = 0; index < count; index++)
                        {
                            TBBUTTON& rTbButton = pTbButtons[index];
                            rTbButton.idCommand = ::m_cuPreviewBar[index];
                            rTbButton.fsStyle = TBSTYLE_BUTTON;
                            rTbButton.fsState = TBSTATE_ENABLED;
                            // 「1ページ/2ページ」ボタンのテキストは読み先を変える
                            UINT nID = rTbButton.idCommand;
                            switch (rTbButton.idCommand)
                            {
                            case AFX_ID_PREVIEW_NUMPAGE:
                                nID = AFX_IDS_ONEPAGE;
                                break;
                            }
                            // ボタンに設定する文字列を取得する
                            LPTSTR source = (LPTSTR)::m_szWindowClass;
                            if (nullptr != source && 
                                0 < ::LoadString(::m_hInst, nID, source, 1 + MAX_PATH))
                            {
                                LPCTSTR result = nullptr;
                                switch (rTbButton.idCommand)
                                {
                                case AFX_ID_PREVIEW_NUMPAGE:
                                    // 「1ページ/2ページ」ボタンのテキストは複製する
                                    result = ::_tcsdup((LPCTSTR)source);
                                    break;
                                default:
                                    // ツールバーのタイトルを取得する
                                    result = ::AfxExtractSubString((LPCTSTR)source, 1);
                                    break;
                                }
                                if (nullptr != result)
                                {
                                    // 印刷プレビューツールバーのボタンに文字列を設定する
                                    rTbButton.iString = (int)::SendMessage(hCtrl, TB_ADDSTRING, 0, (LPARAM)result);
                                    ::free((void*)result);
                                    result = nullptr;
                                }
                            }
                            // コモンコントロールのビットマップの位置を設定する
                            result = FALSE;
                            for (UINT nIndex = 0; FALSE == result && nIndex < nCount; nIndex++)
                            {
                                result = rTbButton.idCommand == LOWORD(::m_cuPreviewBarIcons[nIndex]);
                                if (result)
                                {
                                    rTbButton.iBitmap = offset[0];
                                    if (IDB_HIST_SMALL_COLOR == HIBYTE(HIWORD(::m_cuPreviewBarIcons[nIndex])))
                                    {
                                        rTbButton.iBitmap = offset[1];
                                    }
                                    // ビットマップの位置を設定する
                                    rTbButton.iBitmap += LOBYTE(HIWORD(::m_cuPreviewBarIcons[nIndex]));
                                }
                            };
                        };
                        // ボタンを設定する
                        ::SendMessage(hCtrl, TB_ADDBUTTONS, (WPARAM)count, (LPARAM)pTbButtons);
                        ::SendMessage(hCtrl, TB_AUTOSIZE, 0, 0);
                        result = TRUE;

                        // ボタン設定構造体を解放
                        ::SecureZeroMemory(pTbButtons, value);
                        ::free(pTbButtons);
                        pTbButtons = nullptr;
                    }
                }
            }
            if (result)
            {
                result = FALSE;
                // ステータスバーの設定
                value = WS_CHILD | WS_VISIBLE;
                HWND hCtrl = ::CreateWindow(STATUSCLASSNAME, nullptr, value, 0, 0, 0, 0, hWnd, (HMENU)AFX_IDW_STATUS_BAR, ::m_hInst, nullptr);
                if (::IsWindow(hCtrl))
                {
                    value = 0;
                    HRSRC hRsrc = ::FindResource(::m_hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), (LPTSTR)RT_TOOLBAR);
                    if (hRsrc != nullptr)
                    {
                        HGLOBAL hGlobal = ::LoadResource(::m_hInst, hRsrc);
                        if (hGlobal != nullptr)
                        {
                            LPTOOLBARDATA pData = (LPTOOLBARDATA)::LockResource(hGlobal);
                            if (pData != nullptr && pData->version == 1)
                            {
                                value = pData->count;
                                WORD* pItems = (WORD*)&pData[1];
                                ::m_pStatusBar = (UINT*)::malloc(value * sizeof(UINT));
                                if (nullptr != ::m_pStatusBar)
                                {
                                    for (UINT index = 0; index < value; index++)
                                    {
                                        ::m_pStatusBar[index] = pItems[index];
                                    };
                                }
                                UnlockResource(hGlobal);
                                hGlobal = nullptr;
                            }
                        }
                    }
                    value *= sizeof(UINT);
                    UINT* source = (UINT*)::malloc(value);
                    if (nullptr != source)
                    {
                        ::SecureZeroMemory(source, value);
                        value /= sizeof(UINT);
                        for (UINT index = 0; index < value; index++)
                        {
                            source[index] = 3;
                        };
                        ::SendMessage(hCtrl, SB_SETPARTS, (WPARAM)value, (LPARAM)source);
                        ::free(source);
                        source = nullptr;
                    }
                    // ステータスバーの設定文字列の配列を作成する
                    UINT length = value;
                    length++;
                    length *= sizeof(LPCTSTR);
                    ::m_aStatusBar = (LPCTSTR*)::malloc(length);
                    if (nullptr != ::m_aStatusBar)
                    {
                        ::SecureZeroMemory(::m_aStatusBar, length);
                        // ステータスバーの許可禁止フラグの配列を作成する
                        length = value;
                        length *= sizeof(BOOL);
                        ::m_bStatusBar = (BOOL*)::malloc(length);
                        if (nullptr != ::m_bStatusBar)
                        {
                            ::SecureZeroMemory(::m_bStatusBar, length);

                            // ステータスバーの各ペインの設定文字列の配列を設定する
                            for (UINT index = 0; index < value; index++)
                            {
                                UINT nID = ::m_pStatusBar[index];
                                switch (nID)
                                {
                                case ID_SEPARATOR:
                                    ::m_aStatusBar[index] = ::_tcsdup(_T(""));
                                    break;
                                default:
                                    if (0 < ::LoadString(::m_hInst, nID, (LPTSTR)::m_szWindowClass, 1 + MAX_PATH))
                                    {
                                        LPCTSTR* aArray = nullptr;
                                        ::Split(&aArray, ::m_szWindowClass);
                                        ::m_aStatusBar[index] = ::_tcsdup(aArray[0]);
                                        ::SendMessage(hCtrl, SB_SETTEXT, SBT_OWNERDRAW | index, (LPARAM)FALSE);
                                        ::Split(&aArray);
                                    }
                                    break;
                                }
                            };
                        }
                    }
                    result = TRUE;
                }
            }
            if (result)
            {
                result = FALSE;
                // リストビューの構築
                value = WS_CHILD | WS_VISIBLE;
                value |= LVS_EDITLABELS;
            //  value |= LVS_NOSORTHEADER;
                value |= LVS_REPORT;
                value |= LVS_SINGLESEL;
                value |= LVS_SHAREIMAGELISTS;
                value |= LVS_SHOWSELALWAYS;
                HWND source = ::CreateWindow(WC_LISTVIEW, nullptr, value, 0, 0, 0, 0, hWnd, (HMENU)AFX_IDW_PANE_FIRST, ::m_hInst, nullptr);
                if (::IsWindow(source))
                {
                    // 拡張属性の設定
                    value = (UINT)::SendMessage(source, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0);
                    value |= LVS_EX_FULLROWSELECT;
                    value |= LVS_EX_GRIDLINES;
                    value |= LVS_EX_HEADERDRAGDROP;
                    value |= LVS_EX_INFOTIP;
                    value |= LVS_EX_SUBITEMIMAGES;
                    ::SendMessage(source, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, value);
                    // イメージリスト（大）の設定
                    ::SendMessage(source, LVM_SETIMAGELIST, (WPARAM)LVSIL_NORMAL, (LPARAM)hLargeIcon);
                    // イメージリスト（小）の設定
                    ::SendMessage(source, LVM_SETIMAGELIST, (WPARAM)LVSIL_SMALL, (LPARAM)hSmallIcon);
                    //リストビューのサブクラスを設定
#ifdef _M_IX86
                    ::m_wndList = (WNDPROC)(LONGLONG)::SetWindowLong(source, GWL_WNDPROC, (LONG)(LONGLONG)ListProc);
#else
                    ::m_wndList = (WNDPROC)::SetWindowLongPtr(source, GWLP_WNDPROC, (LONG_PTR)ListProc);
#endif
                    result = TRUE;
                }
            }
            if (result)
            {
                result = FALSE;
                // 検索/置換ダイアログ用構造体の作成
                value = sizeof(FINDREPLACE);
                LPFINDREPLACE source = (LPFINDREPLACE)::malloc(value);
                if (nullptr != source)
                {
                    ::SecureZeroMemory(source, value);

                    source->lStructSize = value;
                    source->hwndOwner   = hWnd;

                    // 検索文字列の用バッファの構築
                    index = 1 + MAX_PATH;
                    value = sizeof(TCHAR) * index;
                    LPTSTR result = (LPTSTR)::malloc(value);
                    if (nullptr != result)
                    {
                        source->wFindWhatLen = (WORD)index;
                        ::SecureZeroMemory(result, value);
                        source->lpstrFindWhat = result;
                    }
                    // 置換文字列の用バッファの構築
                    result = (LPTSTR)::malloc(value);
                    if (nullptr != result)
                    {
                        source->wReplaceWithLen = (WORD)index;
                        ::SecureZeroMemory(result, value);
                        source->lpstrReplaceWith = result;
                    }
                    // 検索/置換ダイアログからのコールバックメッセージを構築
                    WM_FINDREPLACE = RegisterWindowMessage(FINDMSGSTRING);
                    // 検索/置換ダイアログ用構造体を保持
                    ::m_pFindReplace = source;
                }
                result = nullptr != ::m_pFindReplace;
            }
            // メインフォームの構築可否判定
            if (FALSE == result)
            {
                // 失敗した場合、-1を返す
                return -1;
            }
        }
        else
        {
            return ::DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    case WM_DRAWITEM:
        // ステータスバーのオーナー描画
        if (::IsWindow(hWnd) && 0 != lParam)
        {
            // ステータスバーのオーナー描画処理
            LPDRAWITEMSTRUCT source = (LPDRAWITEMSTRUCT)lParam;
            // オーナー描画で設定すするステータスバー文字列を取得する
            LPCTSTR value = ::m_aStatusBar[source->itemID];
            if (nullptr != value)
            {
                // 背景色設定
                ::SetBkColor(source->hDC, GetSysColor(COLOR_MENU));
                // 前景色設定
                UINT index = ::GetSysColor(COLOR_GRAYTEXT);
                if (0 != source->itemData)
                {
                    index = ::GetSysColor(COLOR_WINDOWTEXT);
                }
                ::SetTextColor(source->hDC, index);
                // 透過色設定（不透明）
                ::SetBkMode(source->hDC, OPAQUE);
                // ステータスバーのペインのデバイスコンテキストへ文字列を描画
                ::TextOut(source->hDC, source->rcItem.left, source->rcItem.top + 2, value, (int)::_tcslen(value));
            }
        }
        break;
    case WM_GETMINMAXINFO:
        if (::IsWindow(hWnd))
        {
            // メインフォームのリサイズ時の最小の大きさを設定
            LPMINMAXINFO value = (LPMINMAXINFO)lParam;
            LPSIZE source = &::m_WindowMinSize;
            value->ptMinTrackSize.x = source->cx;
            value->ptMinTrackSize.y = source->cy;
        }
        break;
    case WM_INITIALUPDATE:
        // メインフォームの更新処理
        if (::IsWindow(hWnd))
        {
            HWND source = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
            if (::IsWindow(source))
            {
                // メインフォームのタイトル設定
                {
                    LPCTSTR result = nullptr;
                    int index = 1 + MAX_PATH;
                    size_t value = sizeof(TCHAR) * index;
                    LPTSTR source = (LPTSTR)::malloc(value);
                    if (nullptr != source)
                    {
                        ::SecureZeroMemory(source, value);
                        if (0 < LoadString(::m_hInst, AFX_IDS_UNTITLED, source, index))
                        {
                            if (nullptr != ::m_pOpenFileName &&
                                nullptr != ::m_pOpenFileName->lpstrFileTitle &&
                                NULLSTR != *::m_pOpenFileName->lpstrFileTitle)
                            {
                                ::_tcscpy_s(source, index, ::m_pOpenFileName->lpstrFileTitle);
                            }
                        }
                        result = ::AfxFormatString(AFX_IDS_OBJ_TITLE_INPLACE, ::m_szTitle, (LPCTSTR)source);
                        ::free(source);
                        source = nullptr;
                    }
                    if (nullptr != result)
                    {
                        ::SetWindowText(hWnd, result);
                        ::free((void*)result);
                        result = nullptr;
                    }
                }
                // リストビューの全アイテムの削除
                int iMaxSubItems = ::GetSubItemCount();
                BOOL result = FALSE != ::SendMessage(source, LVM_DELETEALLITEMS, 0, 0);
                for (int iSubItem = iMaxSubItems - 1; result && 0 <= iSubItem; iSubItem--)
                {
                    result = FALSE != ::SendMessage(source, LVM_DELETECOLUMN, iSubItem, 0);
                };
                // データがあるかどうかを判定
                if (result && ::IsOpen())
                {
                    // データがある場合
                    // リストビューのヘッダを構築
                    iMaxSubItems = ::GetFieldCount();
                    UINT value = sizeof(LVCOLUMN);
                    LPLVCOLUMN pLvColumn = (LPLVCOLUMN)::malloc(value);
                    if (nullptr != pLvColumn)
                    {
                        ::SecureZeroMemory(pLvColumn, value);
                        pLvColumn->mask = LVCF_TEXT | LVCF_FMT | LVCF_WIDTH;
                        pLvColumn->fmt  = LVCFMT_LEFT;
                        pLvColumn->cx   = 100;
                        for (int index = 0; result && index < iMaxSubItems; index++)
                        {
                            pLvColumn->pszText = (LPTSTR)::GetFieldName(index);
                            if (nullptr != pLvColumn->pszText)
                            {
                                result = index == (int)SendMessage(source, LVM_INSERTCOLUMN, (WPARAM)index, (LPARAM)pLvColumn);
                                ::free(pLvColumn->pszText);
                                pLvColumn->pszText = nullptr;
                            }
                        };
                        ::free(pLvColumn);
                        pLvColumn = nullptr;
                    }
                    // リストビューのデータ一覧を構築
                    for (::m_iItem = 0; result && !::IsEOF(); m_iItem++)
                    {
                        for (int iSubItem = 0; result && iSubItem < iMaxSubItems; iSubItem++)
                        {
                            result = FALSE;
                            LPCTSTR value = ::GetFieldValue(iSubItem);
                            if (nullptr != value)
                            {
                                result = nullptr != ::DDX_ListViewItemText(source, iSubItem, value);
                                ::free((void*)value);
                                value = nullptr;
                            }
                        };
                        // 次行へ
                        ::MoveNext();
                    };
                }
                // リストビューの先頭を注目先に設定する
                ::m_iItem = ::m_iSubItem = 0;
                // リストビューの許可/禁止処理
                result = 0 < (int)::SendMessage(source, LVM_GETITEMCOUNT, 0, 0);
                ::EnableWindow(source, result);
                if (FALSE != result)
                {
                    // データがある場合、注目位置を選択表示する
                    ::SetRowCol(source);
                }
                // タイマーに関係ない、メニューバーとツールバーの 許可/禁止処理
                HMENU hMenu = ::GetMenu(hWnd);
                UINT value = sizeof(::m_cuEnables) / sizeof(UINT);
                HWND source = ::GetDlgItem(hWnd, AFX_IDW_TOOLBAR);
                for (UINT index = 0; index < value; index++)
                {
                    // 許可禁止判定
                    UINT value = ::m_cuEnables[index];
                    BOOL enable = result;
                    // 「ドキュメントを開く」処理だけ逆転する
                    switch (value)
                    {
                    case ID_FILE_OPEN:
                        enable = FALSE == enable;
                        break;
                    }
                    // メニューバーの許可禁止処理
                    if (::IsMenu(hMenu))
                    {
                        ::EnableMenuItem(hMenu, value, enable ? MF_ENABLED : MF_GRAYED | MF_DISABLED);
                    }
                    // ツールバーの許可禁止処理
                    if (::IsWindow(source))
                    {
                        ::SendMessage(source, TB_ENABLEBUTTON, (WPARAM)value, MAKELPARAM(enable, 0));
                    }
                };
            }
        }
        break;
    case WM_KILLFOCUS:
        // フォーカスが失われようとしている場合の処理
        if (::IsWindow(hWnd))
        {
            HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
            if (::IsWindow(hCtrl) && ::IsWindowVisible(hCtrl))
            {
                ::InvalidateRect(hCtrl, nullptr, TRUE);
            }
        }
        break;
    case WM_MENUSELECT:
        // メニューアイテム上でホバリングする場合の処理
        if (::IsWindow(hWnd))
        {
            // ホバリング中のメニューアイテムのIDを取得
            UINT nID = LOWORD(wParam);
            if (0 != nID)
            {
                // ステータスバーの取得
                HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDW_STATUS_BAR);
                if (::IsWindow(hCtrl))
                {
                    // ステータスバーの表示用文字列の取得
                    int index = 1 + MAX_PATH;
                    LPTSTR value = (LPTSTR)::m_szWindowClass;
                    if (0 < ::LoadString(::m_hInst, nID, value, index))
                    {
                        value = (LPTSTR)::AfxExtractSubString(value, 0);
                        if (nullptr != value)
                        {
                            // ツールバーアイコンの取得
                            HICON hIcon = ::GetToolBarIcon(nID);
                            if (nullptr != hIcon &&
                                INVALID_HANDLE_VALUE != hIcon)
                            {
                                // NOP
                            }
                            else
                            {
                                // アイコンの取得
                                hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                                if (::IsWindow(hCtrl))
                                {
                                    HIMAGELIST hImageList = (HIMAGELIST)::SendMessage(hCtrl, LVM_GETIMAGELIST, LVSIL_SMALL, 0);
                                    if (nullptr != hImageList &&
                                        INVALID_HANDLE_VALUE != hImageList)
                                    {
                                        hIcon = ::ImageList_GetIcon(hImageList, 4, ILD_NORMAL);
                                    }
                                }
                            }
                            // ステータスバーへアイコンを設定する
                            if (nullptr != hIcon &&
                                INVALID_HANDLE_VALUE != hIcon)
                            {
                                index = ::CommandToIndex(ID_SEPARATOR);
                                switch (index)
                                {
                                case IDC_STATIC:
                                    break;
                                default:
                                    ::SendMessage(hCtrl, SB_SETICON, (WPARAM)index, (LPARAM)hIcon);
                                    break;
                                }
                            }
                            // メッセージ文字列をステータスバーへ設定
                            ::SendMessage(hWnd, WM_SETMESSAGESTRING, 0, (LPARAM)value);
                            ::free(value);
                            value = nullptr;
                        }
                    }
                }
            }
        }
        else
        {
            return ::DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    case WM_NOTIFY:
        // 通知処理
        if (0 != lParam && ::IsWindow(hWnd))
        {
            LRESULT result = 0, * pResult = &result;
            LPNMHDR pNMHDR = (LPNMHDR)lParam;
            switch (pNMHDR->code)
            {
            case TTN_NEEDTEXT:
                // ツールバーもしくは印刷プレビュー用ツールバーの
                // ツールチップを表示する
                if (TRUE)
                {
                    LPNMTTDISPINFO pNMTTDispInfo = (LPNMTTDISPINFO)lParam;
                    if (0 == (pNMTTDispInfo->uFlags & TTF_IDISHWND))
                    {
                        UINT nID = (UINT)pNMTTDispInfo->hdr.idFrom;
                        size_t index = (size_t)(1 + MAX_PATH);
                        LPTSTR value = (LPTSTR)::m_szWindowClass;
                        // ツールチップの説明用文字列を取得
                        if (0 < ::LoadString(::m_hInst, nID, value, (int)index))
                        {
                            // ツールバーへ設定されているアイコンを取得
                            HICON hIcon = ::GetToolBarIcon(nID);
                            value = (LPTSTR)::AfxExtractSubString(value, 0);
                            if (nullptr != value)
                            {
                                // ツールチップの表示用文字列を保持する
                                if (nullptr != ::m_sToolTip)
                                {
                                    pNMTTDispInfo->lpszText = ::m_sToolTip;
                                    ::_tcscpy_s(::m_sToolTip, index, value);
                                }
                                if (::IsWindow(::m_hMainWnd))
                                {
                                    HWND hCtrl = nullptr;
                                    HICON hStatusIcon = hIcon;
                                    if (nullptr != hIcon &&
                                        INVALID_HANDLE_VALUE != hIcon)
                                    {
                                        // NOP
                                    }
                                    else
                                    {
                                        // アイコンが無い場合、
                                        // リストビューから取得
                                        hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                                        if (::IsWindow(hCtrl))
                                        {
                                            HIMAGELIST hImageList = (HIMAGELIST)::SendMessage(hCtrl, LVM_GETIMAGELIST, LVSIL_SMALL, 0);
                                            if (nullptr != hImageList)
                                            {
                                                hIcon = ::ImageList_GetIcon(hImageList, 4, ILD_NORMAL);
                                            }
                                        }
                                    }
                                    // ステータスバーを取得
                                    hCtrl = ::GetDlgItem(hWnd, AFX_IDW_STATUS_BAR);
                                    if (::IsWindow(hCtrl))
                                    {
                                        // アイコンを設定するペインを取得
                                        index = ::CommandToIndex(ID_SEPARATOR);
                                        switch (index)
                                        {
                                        case IDC_STATIC:
                                            break;
                                        default:
                                            // ステータスバーへアイコンを設定
                                            ::SendMessage(hCtrl, SB_SETICON, (WPARAM)index, (LPARAM)hStatusIcon);
                                            break;
                                        }
                                    }
                                    // 説明メッセージを設定
                                    ::SendMessage(::m_hMainWnd, WM_SETMESSAGESTRING, 0, (LPARAM)value);
                                }
                                ::free(value);
                                value = nullptr;
                            }
                            // ツールチップへ設定するタイトル文字列を取得する
                            LPTSTR result = (LPTSTR)::m_szWindowClass;
                            value = (LPTSTR)::AfxExtractSubString((LPCTSTR)result, 1);
                            if (nullptr != value)
                            {
                                // メニューバーの該当アイテムに設定されているアクセラレータ文字列を取得する
                                LPCTSTR source = nullptr;
                                // メインフレームのメニューバーへのハンドルを取得
                                HMENU index = ::GetMenu(hWnd);
                                if (::IsMenu(index))
                                {
                                    if (0 < ::GetMenuString(index, nID, result, 1 + MAX_PATH, MF_BYCOMMAND))
                                    {
                                        source = ::AfxExtractSubString(result, 1, _T("\t"));
                                    }
                                }
                                if (nullptr != source)
                                {
                                    // アクセラレータ文字列がある場合、
                                    // タイトル文字列へ追加する
                                    ::_stprintf_s(result, (size_t)(1 + MAX_PATH), _T("%s (%s)"), value, source);
                                    ::free((void*)source);
                                    source = nullptr;
                                }
                                else
                                {
                                    ::_tcscpy_s(result, (size_t)(1 + MAX_PATH), value);
                                }
                                ::free(value);
                                value = nullptr;

                                // ツールチップへアイコンを設定する
                                HICON hToolTipIcon = (HICON)TTI_INFO;
                                if (nullptr != hIcon &&
                                    INVALID_HANDLE_VALUE != hIcon)
                                {
                                    hToolTipIcon = hIcon;
                                }
                                // ツールチップのタイトルを設定
                                ::SendMessage(pNMTTDispInfo->hdr.hwndFrom, TTM_SETTITLE, (WPARAM)hToolTipIcon, (LPARAM)::m_szWindowClass);
                            }
                        }
                    }
                }
                break;
            default:
                // リストビューの各種通知処理
                if (TRUE)
                {
                    // 現在の通知がリストビューのモノかどうかを判定
                    HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                    if (::IsWindow(hCtrl) && hCtrl == pNMHDR->hwndFrom)
                    {
                        switch (pNMHDR->code)
                        {
                        case NM_SETFOCUS:
                            // リストビューがフォーカスを取得した場合の処理
                            if (::IsWindowVisible(hCtrl) &&
                                ::IsWindowEnabled(hCtrl) &&
                                0 < (int)::SendMessage(hCtrl, LVM_GETITEMCOUNT, 0, 0))
                            {
                                ::InvalidateRect(hCtrl, nullptr, TRUE);
                            }
                            break;
                        case NM_KILLFOCUS:
                            // リストビューがフォーカスを失った場合の処理
                            if (::IsWindowVisible(hCtrl))
                            {
                                ::InvalidateRect(hCtrl, nullptr, TRUE);
                            }
                            break;
                        case NM_CLICK:
                            // リストビューをクリックした場合の処理
                            if (TRUE)
                            {
                                LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
                                if (nullptr != pNMItemActivate)
                                {
                                    ::m_iItem = pNMItemActivate->iItem;
                                    ::m_iSubItem = pNMItemActivate->iSubItem;
                                    ::InvalidateRect(hCtrl, nullptr, TRUE);
                                }
                            }
                            break;
                        case LVN_ITEMCHANGED:
                            // リストビューのアイテムを選択変更した場合の処理
                            if (TRUE)
                            {
                                LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
                                if (nullptr != pNMLV)
                                {
                                    UINT value = (UINT)::SendMessage(hCtrl, LVM_GETVIEW, 0, 0);
                                    switch (value)
                                    {
                                    case LVS_REPORT:
                                        break;
                                    default:
                                        ::m_iSubItem = 0;
                                        break;
                                    }
                                    if (::m_iItem != pNMLV->iItem)
                                    {
                                        ::m_iItem = pNMLV->iItem;
                                        ::InvalidateRect(hCtrl, nullptr, TRUE);
                                    }
                                }
                            }
                            break;
                        case LVN_KEYDOWN:
                            // リストビュー上でキーを押した場合の処理
                            if (nullptr != pNMHDR)
                            {
                                LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
                                switch (pLVKeyDow->wVKey)
                                {
                                case VK_LEFT:
                                case VK_RIGHT:
                                    if (TRUE)
                                    {
                                        int iMaxSubItems = ::GetSubItemCount();
                                        if (2 < iMaxSubItems)
                                        {
                                            iMaxSubItems--;
                                            LPLVCOLUMN source = (LPLVCOLUMN)::malloc(sizeof(LVCOLUMN));
                                            if (nullptr != source)
                                            {
                                                ::SecureZeroMemory(source, sizeof(LVCOLUMN));
                                                source->mask = LVCF_ORDER;
                                                if (FALSE != ::SendMessage(hCtrl, LVM_GETCOLUMN, (WPARAM)::m_iSubItem, (LPARAM)source))
                                                {
                                                    int value = source->iOrder;
                                                    switch (pLVKeyDow->wVKey)
                                                    {
                                                    case VK_LEFT:
                                                        if (0 < value)
                                                        {
                                                            value--;
                                                        }
                                                        break;
                                                    case VK_RIGHT:
                                                        if (value < iMaxSubItems)
                                                        {
                                                            value++;
                                                        }
                                                        break;
                                                    }
                                                    if (source->iOrder != value)
                                                    {
                                                        BOOL result = FALSE;
                                                        for (int index = 0; FALSE == result && index <= iMaxSubItems; index++)
                                                        {
                                                            if (FALSE != ::SendMessage(hCtrl, LVM_GETCOLUMN, (WPARAM)index, (LPARAM)source))
                                                            {
                                                                result = value == source->iOrder;
                                                                if (result)
                                                                {
                                                                    ::m_iSubItem = index;
                                                                    ::InvalidateRect(hCtrl, nullptr, TRUE);
                                                                }
                                                            }
                                                        };
                                                    }
                                                }
                                                ::free(source);
                                                source = nullptr;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                            break;
                        case NM_CUSTOMDRAW:
                            // リストビューのカスタムドロー処理
                            if (TRUE)
                            {
                                LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
                                if (nullptr != pNMCD)
                                {
                                    switch (pNMCD->dwDrawStage)
                                    {
                                    case CDDS_PREPAINT:
                                        *pResult = CDRF_NOTIFYITEMDRAW;
                                        break;
                                    case CDDS_ITEMPREPAINT:
                                        *pResult = CDRF_NOTIFYSUBITEMDRAW;
                                        break;
                                    case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
                                        *pResult = CDRF_NEWFONT;
                                        if (TRUE)
                                        {
                                            LPNMLVCUSTOMDRAW pNMLV = reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR);
                                            UINT
                                                value = ::GetWindowLong(hCtrl, GWL_STYLE),
                                                source = 0,
                                                result = 0;
                                            switch (value & LVS_EDITLABELS)
                                            {
                                            case LVS_EDITLABELS:
                                                switch (LVS_EX_FULLROWSELECT & ::SendMessage(hCtrl, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0))
                                                {
                                                case LVS_EX_FULLROWSELECT:
                                                    source = ::GetSysColor(COLOR_WINDOWTEXT);
                                                    result = ::GetSysColor(COLOR_WINDOW);
                                                    if (0 !=  (1 & pNMCD->dwItemSpec))
                                                    {
                                                        // 縞々処理
                                                        result = RGB(204, 255, 255);
                                                    }
                                                    if (::m_iItem != pNMCD->dwItemSpec &&
                                                        ::m_iSubItem == pNMLV->iSubItem)
                                                    {
                                                        switch (value & LVS_SHOWSELALWAYS)
                                                        {
                                                        case LVS_SHOWSELALWAYS:
                                                            result = ::GetSysColor(COLOR_MENU);
                                                            break;
                                                        }
                                                        HWND hFocus = ::GetFocus();
                                                        if (::IsWindow(hFocus) && hCtrl == hFocus)
                                                        {
                                                            // 選択列を
                                                            // 強調表示する
                                                            source = ::GetSysColor(COLOR_HIGHLIGHTTEXT);
                                                            result = ::GetSysColor(COLOR_HIGHLIGHT);
                                                        }
                                                    }
                                                    pNMLV->clrText = source;
                                                    pNMLV->clrTextBk = result;
                                                    break;
                                                }
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            break;
                        case LVN_COLUMNCLICK:
                            // 詳細リストビュー上のカラムのクリック処理
                            if (TRUE)
                            {
                                LPNMLISTVIEW pNMLV =
                                    reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
                                // ソート用構造体の構築
                                UINT value = sizeof(SORTLIST);
                                LPSORTLIST pSortList = (LPSORTLIST)
                                    ::malloc(value);
                                if (nullptr != pSortList &&
                                    ::IsWindow(::m_hMainWnd))
                                {
                                    ::SecureZeroMemory(pSortList, value);
                                
                                    // リストビューのハンドルを取得
                                    pSortList->hWnd = ::GetDlgItem(::m_hMainWnd, AFX_IDW_PANE_FIRST);
                                    if (::IsWindow(pSortList->hWnd))
                                    {
                                        pSortList->iSubItem = pNMLV->iSubItem;
                                        pSortList->bDirection = FALSE;

                                        // リストビューの詳細表示時の
                                        // カラムヘッダフィールド取得用構造体を作成する
                                        value = sizeof(LVCOLUMN);
                                        LPLVCOLUMN pLvColumn = (LPLVCOLUMN)::malloc(value);
                                        if (nullptr != pLvColumn)
                                        {
                                            // リストビューの詳細表示のヘッダに
                                            // クリックしたフィールドに昇順降順インジケータをつける
                                            UINT mask = HDF_SORTDOWN | HDF_SORTUP;
                                            pLvColumn->mask = LVCF_FMT;
                                            int iMaxSubItems = ::GetSubItemCount();
                                            for (int index = 0; index < iMaxSubItems; index++)
                                            {
                                                // カラムヘッダのフィールドを取得
                                                if (FALSE != ::SendMessage(pSortList->hWnd, LVM_GETCOLUMN, (WPARAM)index, (LPARAM)pLvColumn))
                                                {
                                                    // 昇順降順インジケーターの状態を取得
                                                    value = mask;
                                                    value &= pLvColumn->fmt;
                                                    // 昇順降順インジケーターの状態を消去
                                                    pLvColumn->fmt &= ~mask;
                                                    if (pSortList->iSubItem == index)
                                                    {
                                                        switch (value)
                                                        {
                                                        case HDF_SORTUP:
                                                            // 降順設定
                                                            pLvColumn->fmt |= HDF_SORTDOWN;
                                                            pSortList->bDirection = TRUE;
                                                            break;
                                                        default:
                                                            // 昇順設定
                                                            pLvColumn->fmt |= HDF_SORTUP;
                                                            break;
                                                        }
                                                    }
                                                    // カラムヘッダのフィールドを設定
                                                    ::SendMessage(pSortList->hWnd, LVM_SETCOLUMN, (WPARAM)index, (LPARAM)pLvColumn);
                                                }
                                            };
                                            // リストビューのカラムヘッダフィールド取得用構造体の解放
                                            ::free(pLvColumn);
                                            pLvColumn = nullptr;
                                        }
                                        // 現在選択中の行位置を識別するための
                                        // リストビューアイテム取得用構造体を構築
                                        value = sizeof(LVITEM);
                                        LPLVITEM pLvItem = (LPLVITEM)::malloc(value);
                                        if (nullptr != pLvItem)
                                        {
                                            ::SecureZeroMemory(pLvItem, value);
                                            pLvItem->iItem      = ::m_iItem;
                                            pLvItem->iSubItem   = 0;
                                            pLvItem->mask       = LVIF_PARAM;
 
                                            // 現在選択中の行位置を取得する
                                            if (FALSE != ::SendMessage(pSortList->hWnd, LVM_GETITEM, 0, (WPARAM)pLvItem))
                                            {
                                                // ソート開始
                                                if (FALSE != ::SendMessage(pSortList->hWnd, LVM_SORTITEMS, (WPARAM)pSortList, (LPARAM)SortList))
                                                {
                                                    // 選択行位置を検索するための
                                                    // リストビューアイテム検索構造体を構築
                                                    value = sizeof(LVFINDINFO);
                                                    LVFINDINFO* pLvFindInfo = (LVFINDINFO*)::malloc(value);
                                                    if (nullptr != pLvFindInfo)
                                                    {
                                                        ::SecureZeroMemory(pLvFindInfo, value);

                                                        // リストビューアイテムの検索設定
                                                        pLvFindInfo->flags  = LVFI_PARAM;
                                                        pLvFindInfo->lParam = pLvItem->lParam;

                                                        // リストビューを検索する
                                                        int index = (int)::SendMessage(pSortList->hWnd, LVM_FINDITEM, (WPARAM)-1, (LPARAM) pLvFindInfo);
                                                        if (-1 != index)
                                                        {
                                                            // 行の検索位置が発見されたら
                                                            // リストビューの注目位置を変更する
                                                            ::m_iItem = index;
                                                        }

                                                        // リストビュー検索構造体の破棄
                                                        ::SecureZeroMemory(pLvFindInfo, value);
                                                        ::free(pLvFindInfo);
                                                        pLvFindInfo = nullptr;
                                                    }
                                                    // 注目行位置を選択して
                                                    // ステータスバーへ設定する
                                                    ::SetRowCol(pSortList->hWnd);
                                                }
                                            }
                                            // リストビューアイテムの取得用構造体の破棄
                                            value = sizeof(LVITEM);
                                            ::SecureZeroMemory(pLvItem, value);
                                            ::free(pLvItem);
                                            pLvItem = nullptr;
                                        }
                                    }
                                    // リストビューのソート用構造体の破棄
                                    value = sizeof(SORTLIST);
                                    ::SecureZeroMemory(pSortList, value);
                                    ::free(pSortList);
                                    pSortList = nullptr;
                                }
                            }
                            break;
                        case NM_DBLCLK:
                            // リストビュー上でダブルクリックした場合の処理
                            ::SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(ID_OLE_EDIT_PROPERTIES, 0), 0);
                            break;
                        case LVN_BEGINLABELEDIT:
                            // リストビューで編集が開始された場合の処理
                            if (TRUE)
                            {
                                NMLVDISPINFO* pDispInfo = reinterpret_cast<NMLVDISPINFO*>(pNMHDR);
                                HWND hCtrl = pNMHDR->hwndFrom;
                                if (::IsWindow(hCtrl))
                                {
                                    HWND hEdit = (HWND)::SendMessage(hCtrl, LVM_GETEDITCONTROL, 0, 0);
                                    if (::IsWindow(hEdit))
                                    {
                                        LPCTSTR value = ::DDX_ListViewItemText(hCtrl, ::m_iSubItem);
                                        if (nullptr != value)
                                        {
                                            ::SetWindowText(hEdit, value);
                                            ::free((void*)value);
                                            value = nullptr;
                                        }
                                    }
                                }
                                ::MessageBox(IDP_AFXBARRES_TEXT_IS_REQUIRED, MB_ICONINFORMATION);
                                ::InvalidateRect(hCtrl, nullptr, TRUE);
                            }
                            break;
                        case LVN_ENDLABELEDIT:
                            // リストビューで編集が終了した場合の処理
                            if (TRUE)
                            {
                                NMLVDISPINFO* pDispInfo = reinterpret_cast<NMLVDISPINFO*>(pNMHDR);
                                if (nullptr != pDispInfo &&
                                    nullptr != pDispInfo->item.pszText)
                                {
                                    if (NULLSTR != *pDispInfo->item.pszText)
                                    {
                                        // 編集成功したので、
                                        // リストビューへ
                                        // エディットボックスの内容を書き戻す
                                        ::DDX_ListViewItemText(hCtrl, ::m_iSubItem, pDispInfo->item.pszText);
                                        // 編集中状態の設定
                                        ::m_bDirty = TRUE;
                                        // 編集メッセージを表示
                                        ::MessageBox(IDS_EDIT_MENU, MB_ICONINFORMATION);
                                    }
                                    else
                                    {
                                        // バリデーションの例。
                                        // 空文字設定時はエラーを表示し、
                                        // タイマーで再度編集状態にする
                                        ::MessageBox(IDP_AFXBARRES_TEXT_IS_REQUIRED);
                                        ::SetTimer(hWnd, ID_TIMER_REEDIT, 125, nullptr);
                                    }
                                }
                                else
                                {
                                    // キャンセル
                                    ::MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
                                }
                            }
                            break;
                        default:
                            return ::DefWindowProc(hWnd, message, wParam, lParam);
                        }
                    }
                    else
                    {
                        return ::DefWindowProc(hWnd, message, wParam, lParam);
                    }
                }
            }
            return result;
        }
        else
        {
            return ::DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    case WM_PAINT:
        // 描画処理
        if (::IsWindow(hWnd))
        {
            PAINTSTRUCT ps;
            HDC hdc = ::BeginPaint(hWnd, &ps);
            if (nullptr != hdc)
            {
                // TODO: HDC を使用する描画コードをここに追加してください...
                // 印刷プレビュー判定
                HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PREVIEW_BAR);
                if (::IsWindow(hCtrl) && ::IsWindowVisible(hCtrl))
                {
                    // 印刷プレビュー表示
                    OnPrint(hdc, ::m_nPrintCurrentPage);
                }
                ::EndPaint(hWnd, &ps);
            }
        }
        else
        {
            return ::DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    case WM_SETFOCUS:
        // メインフォームがフォーカスを取得した場合の処理
        if (::IsWindow(hWnd))
        {
            HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
            if (::IsWindow(hCtrl) &&
                ::IsWindowVisible(hCtrl) &&
                ::IsWindowEnabled(hCtrl) &&
                0 < (int)::SendMessage(hCtrl, LVM_GETITEMCOUNT, 0, 0))
            {
                ::SetFocus(hCtrl);
            }
        }
        break;
    case WM_SETMESSAGESTRING:
        // ステータスバーへメッセージを表示する場合の処理
        if (::IsWindow(hWnd))
        {
            HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDW_STATUS_BAR);
            if (::IsWindow(hCtrl))
            {
                int result = -1,
                    count = (int)::SendMessage(hCtrl, SB_GETPARTS, 0, 0);
                for (int index = 0; -1 == result && index < count; index++)
                {
                    switch (::m_pStatusBar[index])
                    {
                    case ID_SEPARATOR:
                        result = index;
                        break;
                    }
                };
                if (-1 != result)
                {
                    LPCTSTR value = nullptr;
                    if (0 != wParam)
                    {
                        if (0 < ::LoadString(::m_hInst, (UINT)wParam, (LPTSTR)::m_szWindowClass, 1 + MAX_PATH))
                        {
                            value = ::m_szWindowClass;
                        }
                    }
                    else if (0 != lParam)
                    {
                        value = (LPCTSTR)lParam;
                    }
                    if (nullptr != value)
                    {
                        ::SendMessage(hCtrl, SB_SETTEXT, (WPARAM)result, (LPARAM)value);
                    }
                }
            }
        }
        break;
    case WM_SIZE:
        // メインフォームのサイズ変更処理
        if (::IsWindow(hWnd))
        {
            size_t value = sizeof(RECT);
            LPRECT pRect = (LPRECT)::malloc(value);
            if (nullptr != pRect)
            {
                HWND hCtrl = nullptr;
                pRect->left = 0;
                pRect->top = 0;
                pRect->right = LOWORD(lParam);
                pRect->bottom = HIWORD(lParam);
                LPRECT pRc = (LPRECT)::malloc(value);
                if (nullptr != pRect)
                {
                    // ツールバーの範囲取得処理
                    hCtrl = ::GetDlgItem(hWnd, AFX_IDW_TOOLBAR);
                    if (::IsWindow(hCtrl))
                    {
                        ::SendMessage(hCtrl, message, wParam, lParam);
                        if (::IsWindowVisible(hCtrl))
                        {
                            ::GetWindowRect(hCtrl, pRc);
                            pRect->top += pRc->bottom - pRc->top;
                        }
                    }
                    // 印刷プレビューツールバーの範囲取得処理
                    hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PREVIEW_BAR);
                    if (::IsWindow(hCtrl))
                    {
                        ::SendMessage(hCtrl, message, wParam, lParam);
                        if (::IsWindowVisible(hCtrl))
                        {
                            ::GetWindowRect(hCtrl, pRc);
                            pRect->top += pRc->bottom - pRc->top;
                        }
                    }
                    // ステータスバーの範囲取得処理
                    hCtrl = ::GetDlgItem(hWnd, AFX_IDW_STATUS_BAR);
                    if (::IsWindow(hCtrl))
                    {
                        ::SendMessage(hCtrl, message, wParam, lParam);
                        if (::IsWindowVisible(hCtrl))
                        {
                            ::GetWindowRect(hCtrl, pRc);
                            pRect->bottom -= pRc->bottom - pRc->top;
                        }
                    }
                    ::free(pRc);
                    pRc = nullptr;
                }
                // リストビューのリサイズ処理
                hCtrl = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                if (::IsWindow(hCtrl))
                {
                    ::MoveWindow(hCtrl, pRect->left, pRect->top, pRect->right - pRect->left, pRect->bottom - pRect->top, TRUE);
                }
                // ステータスバーのペインのリサイズ処理
                hCtrl = ::GetDlgItem(hWnd, AFX_IDW_STATUS_BAR);
                if (::IsWindow(hCtrl))
                {
                    // ペインの個数を取得
                    value = sizeof(UINT) * (size_t)::SendMessage(hCtrl, SB_GETPARTS, 0, 0);
                    // ペイン位置の配列を作成
                    UINT* source = (UINT*)::malloc(value);
                    if (nullptr != source)
                    {
                        ::SecureZeroMemory(source, value);
                        int sum = 0;
                        HDC hDC = ::GetDC(hWnd);
                        if (nullptr != hDC && INVALID_HANDLE_VALUE != hDC)
                        {
                            LPSIZE pSize = (LPSIZE)::malloc(sizeof(SIZE));
                            if (nullptr != pSize)
                            {
                                ::SecureZeroMemory(pSize, sizeof(SIZE));
                                // ペイン個数ループ
                                value /= sizeof(UINT);
                                for (UINT index = 0; index < value; index++)
                                {
                                    LPCTSTR pValue = nullptr;
                                    switch (::m_pStatusBar[index])
                                    {
                                    case ID_SEPARATOR:
                                        break;
                                    default:
                                        // ペインへ設定している文字列の幅を取得
                                        pValue = ::m_aStatusBar[index];
                                        if (nullptr != pValue &&
                                            // 各文字列の幅を取得
                                            ::GetTextExtentPoint32(hDC, pValue, (int)::_tcslen(pValue), pSize))
                                        {
                                            source[index] = pSize->cx;
                                        }
                                        sum += source[index];
                                        break;
                                    }
                                };
                                ::free(pSize);
                                pSize = nullptr;
                            }
                            ::ReleaseDC(hWnd, hDC);
                        }
                        sum += 17;// ←ステータスバーのグリップの幅
                        int expand = (pRect->right - pRect->left);
                        if (sum < expand)
                        {
                            expand -= sum;
                        }
                        else
                        {
                            expand = 0;
                        }
                        // 各ペインの幅を設定する
                        sum = 0;
                        for (UINT index = 0; index < value; index++)
                        {
                            switch (::m_pStatusBar[index])
                            {
                            case ID_SEPARATOR:
                                sum += expand;
                                break;
                            default:
                                sum += source[index];
                                break;
                            }
                            source[index] = sum;
                        };
                        // 各ペインnの幅を再設定する
                        ::SendMessage(hCtrl, SB_SETPARTS, (WPARAM)value, (LPARAM)source);
                        ::free(source);
                        source = nullptr;
                    }
                }
                // リストビューの移動範囲構造体の破棄
                ::free(pRect);
                pRect = nullptr;
            }
        }
        else
        {
            return ::DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    case WM_TIMER:
        switch (wParam)
        {
        case ID_TIMER_INTERVAL:
            // タイマー処理
            // 現在時刻の取得
            ::GetLocalTime(&::m_systemTime);
            if (::IsWindow(hWnd))
            {
                HMENU hMenu = ::GetMenu(hWnd);
                BOOL result = nullptr != ::m_pOpenFileName &&
                    nullptr != ::m_pOpenFileName->lpstrFile &&
                    NULLSTR != *::m_pOpenFileName->lpstrFile;
                HWND
                    hCtrl = ::GetDlgItem(hWnd, AFX_IDW_TOOLBAR),
                    hList = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST),
                    hEdit = (HWND)::SendMessage(hList, LVM_GETEDITCONTROL, 0, 0);
                // 設定変更するメニューバーアイテムとツールバーボタンの
                // IDの一覧
                UINT value = sizeof(::m_cuEnablesTimer) / sizeof(UINT);
                for (UINT index = 0; index < value; index++)
                {
                    UINT source = ::m_cuEnablesTimer[index];
                    // 許可/禁止判定
                    BOOL enable = FALSE;
                    if (result)
                    {
                        if (::IsWindow(hEdit))
                        {
                            int enter = -1, leave = -1;
                            switch (source)
                            {
                            case ID_EDIT_COPY:
                            case ID_EDIT_CUT:
                                // 「コピー」「切り取り」処理の許可禁止判定
                                ::SendMessage(hEdit, EM_GETSEL, (WPARAM)&enter, (LPARAM)&leave);
                                enable = -1 != enter && -1 != leave && enter != leave;
                                break;
                            case ID_EDIT_PASTE:
                                // 「貼り付け」処理の許可禁止判定
#ifdef _UNICODE
                                enable = ::IsClipboardFormatAvailable(CF_UNICODETEXT);
#else
                                enable = ::IsClipboardFormatAvailable(CF_TEXT);
#endif
                                break;
                            case ID_EDIT_SELECT_ALL:
                                // 「すべて選択」処理の許可禁止判定
                                enable = TRUE;
                                break;
                            case ID_EDIT_UNDO:
                                // 「元に戻す」処理の許可禁止判定
                                enable = FALSE != ::SendMessage(hEdit, EM_CANUNDO, 0, 0);
                                break;
                            }
                        }
                    }
                    // メニューバーの許可/禁止設定
                    if (::IsMenu(hMenu))
                    {
                        ::EnableMenuItem(hMenu,
                            source, enable ?
                            MF_ENABLED :
                            MF_GRAYED | MF_DISABLED);
                    }
                    // ツールバーの許可禁止/設定
                    if (::IsWindow(hCtrl))
                    {
                        ::SendMessage(hCtrl, TB_ENABLEBUTTON,
                            (WPARAM)source, MAKELPARAM(enable, 0));
                    }
                };
                // ステータスバーの設定
                hCtrl = ::GetDlgItem(hWnd, AFX_IDW_STATUS_BAR);
                if (::IsWindow(hCtrl))
                {
                    UINT count = (UINT)::SendMessage(hCtrl, SB_GETPARTS, 0, 0);
                    for (UINT index = 0; index < count; index++)
                    {

                        UINT position = 0, nID = ::m_pStatusBar[index];
                        switch (nID)
                        {
                        case ID_SEPARATOR:
                            break;
                        default:
                            ::_tcscpy_s((LPTSTR)::m_szWindowClass, (size_t)(1 + MAX_PATH), ::m_aStatusBar[index]);
                            result = ::m_bStatusBar[index];
                            switch (nID)
                            {
                            case ID_INDICATOR_CAPS:
                                // CAPSの許可禁止処理
                                result = 0 != (1 & ::GetKeyState(VK_CAPITAL));
                                break;
                            case ID_INDICATOR_NUM:
                                // NUMの許可禁止処理
                                result = 0 != (1 & ::GetKeyState(VK_NUMLOCK));
                                break;
                            case ID_INDICATOR_SCRL:
                                // SCRLの許可禁止処理
                                result = 0 != (1 & ::GetKeyState(VK_SCROLL));
                                break;
                            case ID_INDICATOR_ROW:
                                // 行位置の表示処理
                                result = 0 < (int)::SendMessage(hList, LVM_GETITEMCOUNT, 0, 0);
                                if (result)
                                {
                                    position = 1 + ::m_iItem;
                                }
                                ::_stprintf_s((LPTSTR)::m_szWindowClass, (size_t)(1 + MAX_PATH), _T("行:% 5d"), position);
                                break;
                            case ID_INDICATOR_COL:
                                // 列位置の表示処理
                                result = 0 < (int)::SendMessage(hList, LVM_GETITEMCOUNT, 0, 0);
                                if (result)
                                {
                                    position = 1 + ::m_iSubItem;
                                }
                                ::_stprintf_s((LPTSTR)::m_szWindowClass, (size_t)(1 + MAX_PATH), _T("列:% 3d"), position);
                                break;
                            case ID_INDICATOR_MODIFY:
                                // 編集フラグ処理
                                result = ::m_bDirty;
                                break;
                            case ID_INDICATOR_DATE:
                                // 現在日付処理
                                result = TRUE;
                                ::_stprintf_s((LPTSTR)::m_szWindowClass, (size_t)(1 + MAX_PATH), _T("%04d/%02d/%02d"), ::m_systemTime.wYear, ::m_systemTime.wMonth, ::m_systemTime.wDay);
                                break;
                            case ID_INDICATOR_TIME:
                                // 現在時刻処理
                                result = TRUE;
                                ::_stprintf_s((LPTSTR)::m_szWindowClass, (size_t)(1 + MAX_PATH), ::m_systemTime.wMilliseconds < 500 ? _T("%02d:%02d") : _T("%02d %02d"), ::m_systemTime.wHour, ::m_systemTime.wMinute);
                                break;
                            }
                            // ステータスバーのペインが変更された場合、
                            // オーナー描画する
                            if (::IsWindow(hWnd) &&
                                0 == ::_tcscmp(::m_aStatusBar[index], ::m_szWindowClass) &&
                                result == ::m_bStatusBar[index])
                            {
                                // ペインが更新されていない（NOP）
                            }
                            else
                            {
                                // ステータスバーの各ペインの許可禁止を更新
                                ::m_bStatusBar[index] = result;
                                // ステータスバーの各ペインの文字列を更新
                                ::_tcscpy_s((LPTSTR)::m_aStatusBar[index], 1 + ::_tcslen(::m_aStatusBar[index]), ::m_szWindowClass);
                                // ステータスバーをオーナー描画する
                                ::SendMessage(hCtrl, SB_SETTEXT, SBT_OWNERDRAW | index, (LPARAM)result);
                            }
                            break;
                        }
                    };
                }
            }
            break;
        case ID_TIMER_REEDIT:
            // リストビューの編集の再起動処理
            ::KillTimer(hWnd, wParam);
            ::SendMessage(hWnd, WM_COMMAND, MAKEWPARAM(ID_OLE_EDIT_PROPERTIES, 0), 0);
            break;
        }
        break;
    case WM_DESTROY:
        ::PostQuitMessage(0);
        break;
    default:
        return ::DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// バージョン情報ボックスのメッセージ ハンドラーです。
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        if (TRUE)
        {
            size_t value = (size_t)(1 + MAX_PATH);
            LPTSTR result = (LPTSTR)::m_szWindowClass;
            if (0 < ::GetWindowText(hDlg, result, (int)value))
            {
                LPCTSTR source = ::_tcsdup(result);
                if (nullptr != source)
                {
                    ::_tcscpy_s(result, value, ::m_szTitle);
                    ::_tcscat_s(result, value, source);
                    ::free((void*)source);
                    source = nullptr;
                }
                ::SetWindowText(hDlg, result);
            }
            if (0 < ::GetModuleFileName(nullptr, result, (DWORD)value))
            {
                DWORD
                    uDummy = 0,
                    uSize = ::GetFileVersionInfoSize(result, &uDummy);
                if (0 < uSize)
                {
                    BYTE* source = (BYTE*)::malloc(uSize);
                    if (nullptr != source)
                    {
                        if (::GetFileVersionInfo(result, 0, uSize, source))
                        {
                            LPVOID pValue = nullptr;
                            UINT count = (UINT)(sizeof(::m_pAboutValueKeys) / sizeof(LPCTSTR));
                            for (UINT index = 0; index < count; index++)
                            {
                                ::_tcscpy_s(result, value, _T("\\StringFileInfo\\041104b0\\"));
                                ::_tcscat_s(result, value, ::m_pAboutValueKeys[index]);
                                uDummy = 0;
                                if (VerQueryValue(source, result, &pValue, (PUINT)&uDummy))
                                {
                                    HWND hWnd = ::GetDlgItem(hDlg, IDC_STATIC1 + index);
                                    if (::IsWindow(hWnd))
                                    {
                                        ::GetWindowText(hWnd, result, (int)value);
                                        ::_tcscat_s(result, value, (LPCTSTR)pValue);
                                        ::SetWindowText(hWnd, result);
                                    }
                                }
                            };
                        }
                        ::free(source);
                        source = nullptr;
                    }
                }
            }
        }
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDOK:
        case IDCANCEL:
            ::EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}


UINT_PTR CALLBACK PrintSetupProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    // 印刷ダイアログの処理、
    // 特に何もしていない
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;
    }
    return (UINT_PTR)FALSE;
}


INT_PTR CALLBACK    PrintProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    // 印刷中dファイアログの処理
    switch (message)
    {
    case WM_INITDIALOG:
        // ステータスバーへ印刷中のデバイス名を表示するための文字列
        if (nullptr != ::m_pPrinterName)
        {
            ::free((void*)::m_pPrinterName);
            ::m_pPrinterName = nullptr;
        }
        ::m_hPrinter = nullptr;
        // プリンタの設定状態を取得
        if (nullptr != ::m_pOpenFileName &&
            nullptr != ::m_pOpenFileName->lpstrFileTitle &&
            nullptr != ::m_pPrintDlg &&
            nullptr != ::m_pPrintDlg->hDevMode &&
            INVALID_HANDLE_VALUE != ::m_pPrintDlg->hDevMode &&
            nullptr != ::m_pPrintDlg->hDC &&
            INVALID_HANDLE_VALUE != ::m_pPrintDlg->hDC)
        {
            // ページ番号の取得（全体）
            ::m_nPrintCurrentPage = ::m_pPrintDlg->nMinPage;
            ::m_nPrintMaxPage = ::m_pPrintDlg->nMaxPage;
            switch ((PD_SELECTION | PD_ALLPAGES) & m_pPrintDlg->Flags)
            {
            case PD_SELECTION:
                // ページ番号の取得（印刷ページ番号の取得）
                ::m_nPrintCurrentPage = ::m_pPrintDlg->nFromPage;
                ::m_nPrintMaxPage = ::m_pPrintDlg->nToPage;
                break;
            }
            BOOL result = FALSE;
            // 印刷ダイアログ選択をプリンタデバイスへ書き戻す
            LPCTSTR value = SetPrinter(::m_pPrintDlg->hDevMode, hDlg);
            if (nullptr == value)
            {
                result = FALSE;
                // StartDoc() 用ドキュメント構造体を作成
                LPDOCINFO pDocInfo = (LPDOCINFO)::malloc(sizeof(DOCINFO));
                if (nullptr != pDocInfo)
                {
                    ::SecureZeroMemory(pDocInfo, sizeof(DOCINFO));
                    pDocInfo->lpszDocName = ::m_pOpenFileName->lpstrFileTitle;
#ifdef DEF_PRINTER
                    result = 0 < ::StartDoc(::m_pPrintDlg->hDC, pDocInfo);
#else
                    result = TRUE;
#endif
                    ::free(pDocInfo);
                    pDocInfo = nullptr;
                }
                if (result)
                {
                    // 印刷用にタイマーを起動。
                    ::SetTimer(hDlg, ID_TIMER_INTERVAL, 1000, nullptr);
                }
            }
            else
            {
                ::MessageBox(value);
            }
        }
        else
        {
            ::SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
        }
        break;
    case WM_TIMER:
        // タイマーで順次印刷
        switch ((UINT)wParam)
        {
        case ID_TIMER_INTERVAL:
            if (::IsWindow(hDlg) &&
                nullptr != ::m_hPrinter && 
                nullptr != ::m_pPrintDlg &&
                nullptr != ::m_pPrintDlg->hDC &&
                INVALID_HANDLE_VALUE != ::m_pPrintDlg->hDC)
            {
                // 印刷中ダイアログボックスの「ページ」ラベルの取得
                HWND hCtrl = ::GetDlgItem(hDlg, AFX_IDC_PRINT_PAGENUM);
                if (nullptr != hCtrl && ::IsWindow(hCtrl))
                {
                    // 印刷中のページ番号の作成
                    ::_stprintf_s((LPTSTR)::m_szWindowClass, (size_t)(1 + MAX_PATH), _T("%d / %d"), ::m_nPrintCurrentPage, ::m_nPrintMaxPage);
                    // 「ページ」ラベルへ書き込み
                    ::SetWindowText(hCtrl, ::m_szWindowClass);
                    if (nullptr != ::m_pPrinterName)
                    {
                        // 一旦保護
                        LPCTSTR value = ::_tcsdup(::m_szWindowClass);
                        if (nullptr != value)
                        {
                            // 「%1へ出力中」をリソースから読み出し
                            LPCTSTR source = ::AfxFormatString(AFX_IDS_PRINTONPORT, ::m_pPrinterName);
                            if (nullptr != source)
                            {
                                // 作成した文字列を作業バッファへコピー
                                ::_tcscpy_s((LPTSTR)::m_szWindowClass, (size_t)(1 + MAX_PATH), source);
                                // 保護したページ番号を追記
                                ::_tcscat_s((LPTSTR)::m_szWindowClass, (size_t)(1 + MAX_PATH), value);
                                // ステータスバーへ表示
                                ::MessageBox(::m_szWindowClass, MB_ICONINFORMATION);
                                ::free((void*)source);
                                source = nullptr;
                            }
                            ::free((void*)value);
                            value = nullptr;
                        }
                    }
                }

                BOOL result = FALSE;
                HDC hDC = ::m_pPrintDlg->hDC;
                if (FALSE)// 印刷中判定
                {
                    // 印刷中。続行
                }
                else
                {
                    if (::m_nPrintCurrentPage <= ::m_nPrintMaxPage)
                    {
#ifdef DEF_PRINTER
                        ::StartPage(hDC);
#endif
                        result = OnPrint(hDC, ::m_nPrintCurrentPage);
                        if (result)
                        {
#ifdef DEF_PRINTER
                            ::EndPage(hDC);
#endif
                            ::m_nPrintCurrentPage++;
                        }
                        if (result)
                        {
#ifdef DEF_PRINTER
                            ::EndDoc(hDC);
#endif
                            if (::m_nPrintCurrentPage <= ::m_nPrintMaxPage)
                            {
                                // 処理続行
                            }
                            else
                            {
                                // 正常終了
                                ::SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDOK, 0), 0);
                            }
                        }
                        else
                        {
                            // 中断処理は？
                            ::AbortPrinter(::m_hPrinter);
                            ::SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
                        }
                    }
                    else
                    {
                        // 範囲外
                        ::SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0);
                    }
                }
            }
            break;
        }
        return TRUE;
    case WM_COMMAND:
        // 印刷終了処理
        switch (LOWORD(wParam))
        {
        case IDOK:
        case IDCANCEL:
            // タイマー停止
            ::KillTimer(hDlg, ID_TIMER_INTERVAL);
            // プリンタハンドルを閉じる
            if (nullptr != ::m_hPrinter)
            {
                ::ClosePrinter(::m_hPrinter);
                ::m_hPrinter = nullptr;
            }
            // プリンタ名を破棄
            if (nullptr != ::m_pPrinterName)
        {
                ::free((void*)::m_pPrinterName);
                ::m_pPrinterName = nullptr;
            }
            // 注目位置を先頭へ移動
            ::m_iItem = ::m_iSubItem = 0;
            // 「印刷中」ダイアログを終了する
            ::EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}


LRESULT CALLBACK    ListProc(HWND hWnd, UINT message,
    WPARAM wParam, LPARAM lParam)
{
    LRESULT result = 0;
    // リストビューの処理
    switch (message)
    {
    case WM_COMMAND:
        // コマンド処理を親へ渡す
        if (::IsWindow(::m_hMainWnd))
        {
            ::SendMessage(::m_hMainWnd, message, wParam, lParam);
        }
        break;
    case WM_PAINT:
        // 再描画しようとする場合、
        // フローティングエディタが有効であれば、所定の位置へ移動する
        if (TRUE)
        {
            // フローティングエディタを取得する
            HWND hEdit = (HWND)::SendMessage(hWnd, LVM_GETEDITCONTROL, 0, 0);
            if (::IsWindow(hEdit))
            {
                size_t value = sizeof(RECT);
                LPRECT pRc = (LPRECT)::malloc(value);
                if (nullptr != pRc)
                {
                    // 所定の位置を取得
                    pRc->left = LVIR_LABEL;
                    pRc->top = ::m_iSubItem;
                    if (FALSE != ::SendMessage(hWnd, LVM_GETSUBITEMRECT,
                        (WPARAM)m_iItem, (LPARAM)pRc))
                    {
                        // フローティングエディタを移動
                        ::MoveWindow(hEdit, pRc->left, pRc->top, pRc->right - pRc->left, pRc->bottom - pRc->top, TRUE);
                    }
                    ::free(pRc);
                    pRc = nullptr;
                }
            }
        }
        break;
    }
    return ::CallWindowProc(::m_wndList, hWnd, message, wParam, lParam);
}


int CALLBACK SortList(LPARAM lParam1, LPARAM lParam2, LPARAM lParam3)
{
    // リストビューを並び替える
    int result = 0;
    if (0 != lParam3)
    {
        // パラメータ3を並べ替え用構造体へ展開する
        LPSORTLIST source = (LPSORTLIST)lParam3;
        if (::IsWindow(source->hWnd))
        {
            // リストビューの検索構造体の作成
            UINT value = sizeof(LVFINDINFO);
            LVFINDINFO* pLvFindInfo = (LVFINDINFO*)::malloc(value);
            if (nullptr != pLvFindInfo)
            {
                ::SecureZeroMemory(pLvFindInfo, value);
                pLvFindInfo->flags = LVFI_PARAM;

                // ２項目を取得する
                // 検索キーを作成
                LPARAM params[] = { lParam1,lParam2 };
                // 検索結果に対応する文字列の取得用配列
                LPCTSTR results[] = { nullptr,nullptr };
                for (int index = 0; index < 2; index++)
                {
                    pLvFindInfo->lParam = params[index];
                    // 該当行位置の検索
                    int iItem = (int)::SendMessage(source->hWnd, LVM_FINDITEM, (WPARAM)-1, (LPARAM)pLvFindInfo);
                    if (-1 != iItem)
                    {
                        // 該当行が取得できたので、該当セルの文字列を取得する
                        ::m_iItem = iItem;
                        results[index] = ::DDX_ListViewItemText(source->hWnd, source->iSubItem);
                    }
                };

                // 大小関係の判定
                if (nullptr != results[0] && nullptr != results[1])
                {
                    result = ::_tcscmp(results[0], results[1]);
                }
                if (nullptr != results[1])
                {
                    // 文字列２の破棄
                    ::free((void*)results[1]);
                    results[1] = nullptr;
                }
                if (nullptr != results[0])
                {
                    // 文字列１の破棄
                    ::free((void*)results[0]);
                    results[0] = nullptr;
                }

                // リストビューの検索構造体の破棄
                ::SecureZeroMemory(pLvFindInfo, value);
                ::free(pLvFindInfo);
                pLvFindInfo = nullptr;
            }
        }
        // 並べ替えの方向を反映する
        if (source->bDirection)
        {
            result = -result;
        }
    }
    return result;
}


UINT    Split(LPCTSTR** results, LPCTSTR source, LPCTSTR delimiter)
{
    // 文字分割
    UINT result = 0;
    if (nullptr != results)
    {
        // 分割文字列の破棄
        if (nullptr != *results)
        {
            if (nullptr != **results)
            {
                ::free((void*)**results);
                **results = nullptr;
            }
            ::free((void*)*results);
            *results = nullptr;
        }
        // 引数判定
        if (nullptr != source &&
            nullptr != delimiter &&
            NULLSTR != *delimiter)
        {
            // 分割後の文字配列の要素数の取得
            size_t
                index = 2,
                length = ::_tcslen(delimiter);
            LPCTSTR
                enter = source,
                leave = ::_tcsstr(enter, delimiter);
            while (nullptr != leave)
            {
                index++;
                enter = &leave[length];
                leave = ::_tcsstr(enter, delimiter);
            };
            // 文字配列の作成
            LPCTSTR* value = (LPCTSTR*)::malloc(index * sizeof(LPCTSTR));
            if (nullptr != value)
            {
                ::SecureZeroMemory(value, index * sizeof(LPCTSTR));
                // 各配列へ文字列を設定
                enter = ::_tcsdup(source);
                leave = ::_tcsstr(enter, delimiter);
                while (nullptr != leave)
                {
                    value[result++] = enter;
                    enter = &leave[length];
                    *((LPTSTR)leave) = NULLSTR;
                    leave = ::_tcsstr(enter, delimiter);
                };
                // デリミタ数＋１＝最後の要素を配列へ設定する
                value[result++] = enter;
                *results = value;
            }
        }
    }
    return result;
}


LPCTSTR replace(LPCTSTR source, LPCTSTR pOld, LPCTSTR pNew, UINT* pCount)
{
    // 文字置換
    int count = 0;
    LPCTSTR result = nullptr;
    if (nullptr != source)
    {
        if (nullptr != pOld && NULLSTR != *pOld && nullptr != pNew)
        {
            int nNew = (int)::_tcslen(pNew),
                index = (int)::_tcslen(pOld),
                value = nNew - index;
            // 更新個数を取得
            LPCTSTR
                enter = source,
                leave = ::_tcsstr(enter, pOld);
            while (nullptr != leave)
            {
                count++;
                enter = &leave[index];
                leave = ::_tcsstr(enter, pOld);
            };
            // 更新後のバッファサイズを取得
            value *= count;
            value += (int)::_tcslen(source);
            value++;
            // 置換後の文字列バッファを作成
            result = (LPCTSTR)::malloc(sizeof(TCHAR) * value);
            if (nullptr != result)
            {
                // 置換後の文字列バッファを初期化
                ::SecureZeroMemory((LPTSTR)result, sizeof(TCHAR) * value);
                // 各々をコピー
                LPCTSTR
                    // 注目位置を指定文字列の先頭に再設定
                    enter = source,
                    // 文字列を検索
                    leave = ::_tcsstr(enter, pOld);
                while (nullptr != leave)
                {
                    // 注目位置から発見位置までを追記
                    ::_tcsncat_s((LPTSTR)result, (size_t)value,
                        enter, leave - enter);
                    // 末尾に更新後の文字列を追記
                    ::_tcscat_s((LPTSTR)result, (size_t)value, pNew);
                    // 注目位置を更新
                    enter = &leave[index];
                    // 文字列を検索
                    leave = ::_tcsstr(enter, pOld);
                };
                // 注目位置から残りの文字列を追記
                ::_tcscat_s((LPTSTR)result, (size_t)value, enter);
            }
        }
        else
        {
            // 更新できない場合、更新前文字列をコピーして返す
            result = ::_tcsdup(source);
        }
    }
    if (nullptr != pCount)
    {
        // 更新個数を返す
        *pCount = (UINT)count;
    }
    return result;
}


LPCTSTR AfxExtractSubString(LPCTSTR source, int position, LPCTSTR delimiter)
{
    LPTSTR result = nullptr;
    // 引数確認
    if (nullptr != source && nullptr != delimiter && NULLSTR != *delimiter)
    {
        int count = 0;
        size_t
            length = ::_tcslen(delimiter);
        // デリミタを数える
        LPCTSTR
            enter = source,
            leave = ::_tcsstr(enter, delimiter);
        while (nullptr != leave)
        {
            count++;
            enter = &leave[length];
            leave = ::_tcsstr(enter, delimiter);
        };
        // デリミタ数＋１で要素数になる
        count++;
        // 指定項目数が要素数よりも小さい場合、部分文字列を取得する
        if (position < count)
        {
            enter = source;
            leave = ::_tcsstr(enter, delimiter);
            int index = 0;
            for (; nullptr == result && nullptr != leave && index < count; index++)
            {
                // 指定位置
                if (index == position)
                {
                    // 指定位置の文字長を取得
                    size_t size = leave - enter;
                    // 文字長からバッファ長さを作成
                    length = sizeof(TCHAR) * (size + 1);
                    // 文字列バッファを構築
                    result = (LPTSTR)::malloc(length);
                    if (nullptr != result)
                    {
                        ::SecureZeroMemory((void*)result, length);
                        // 構築したバッファへ部分文字列をコピーする
                        ::_tcsncpy_s(result, size + 1, enter, size);
                    }
                }
                // 次の分割項目へ移動
                enter = &leave[length];
                leave = ::_tcsstr(enter, delimiter);
            };
            // 分割項目が見つからない場合現在位置を返す。（最後の分割分の文字列）
            if (nullptr == result && index == position && nullptr != enter)
            {
                // 最後の分割文字列
                result = ::_tcsdup(enter);
            }
        }
    }
    return (LPCTSTR)result;
}


LPCTSTR AfxFormatString(UINT nID, LPCTSTR arg1, LPCTSTR arg2)
{
    LPCTSTR result = nullptr;
    LPTSTR value = (LPTSTR)::m_szWindowClass;
    if (nullptr != value && nullptr != arg1)
    {
        // ID に該当する文字列を取得
        if (0 < ::LoadString(::m_hInst, nID, value, 1 + MAX_PATH))
        {
            // arg1 がある場合、%1 を arg1 で置き換え
            result = replace(value, _T("%1"), arg1);
            // arg2 があるかどうかを判定
            if (nullptr != arg2 && nullptr != result)
            {
                // 戻り値を入力値に置き換えて置換する
                value = (LPTSTR)result;

                // arg2 がある場合、%2 を arg2 で置き換え
                result = ::replace(value, _T("%2"), arg2);

                // 置き換えた入力値を解放
                ::free(value);
                value = nullptr;
            }
        }
    }
    return result;
}


UINT MessageBox(UINT nID, UINT style)
{
    UINT result = IDOK;
    int index = 1 + MAX_PATH;
    UINT value = sizeof(TCHAR) * index;
    LPTSTR source = (LPTSTR)::malloc(value);
    if (nullptr != source)
    {
        ::SecureZeroMemory(source, value);
        if (0 < ::LoadString(::m_hInst, nID, source, index))
        {
            result = MessageBox(source, style);
        }
        ::SecureZeroMemory(source, value);
        ::free((void*)source);
        source = nullptr;
    }
    return result;
}


UINT MessageBox(LPCTSTR caption, UINT style)
{
    UINT value = MB_ICONINFORMATION | MB_ICONEXCLAMATION;
    value &= style;
    value >>= 4;
    if (4 < value)
    {
        value = 0;
    }
    UINT result = IDOK;
    HWND hWnd = ::m_hMainWnd;
    if (::IsWindow(hWnd))
    {
        int index = ::CommandToIndex(ID_SEPARATOR);
        switch (index)
        {
        case IDC_STATIC:
            break;
        default:
            if (TRUE)
            {
                // アイコンを取得
                HICON source = nullptr;
                // リストビューを取得
                HWND result = ::GetDlgItem(hWnd, AFX_IDW_PANE_FIRST);
                if (::IsWindow(result))
                {
                    // ステータス用イメージリストを取得
                    HIMAGELIST index = (HIMAGELIST)::SendMessage(result, LVM_GETIMAGELIST, LVSIL_SMALL, 0);
                    if (nullptr != index && INVALID_HANDLE_VALUE != index)
                    {
                        // アイコンを取得
                        source = ::ImageList_GetIcon(index, value, ILD_NORMAL);
                    }
                }
                // アイコンの取得判定
                if (nullptr != source && INVALID_HANDLE_VALUE != source)
                {
                    // ステータスバーのハンドルを取得
                    result = ::GetDlgItem(hWnd, AFX_IDW_STATUS_BAR);
                    if (::IsWindow(result))
                    {
                        ::SendMessage(result, SB_SETICON, (WPARAM)index, (LPARAM)source);
                    }
                }
            }
        }
        // ステータスバーへキャプションを設定する
        ::SendMessage(hWnd, WM_SETMESSAGESTRING, 0, (LPARAM)caption);
        index = 1 + MAX_PATH;
        size_t length = sizeof(TCHAR) * index;
        LPTSTR source = nullptr;
        switch (style)
        {
        case MB_ICONINFORMATION:
            break;
        default:
            // スタイルに対応するタイトル用文字列バッファを作成
            source = (LPTSTR)::malloc(length);
            if (nullptr != source)
            {
                ::SecureZeroMemory(source, length);

                // タイトルの候補一覧をリソースから取得する
                if (0 < ::LoadString(::m_hInst, IDS_EDIT_TITLE, source, index))
                {
                    LPCTSTR* aArray = nullptr;
                    ::Split(&aArray, source);
                    // メッセージボックスを起動
                    result = ::MessageBox(hWnd, caption, aArray[value], style);
                    ::Split(&aArray);
                }
                // タイトル用文字列バッファを解放
                ::SecureZeroMemory(source, length);
                ::free((void*)source);
                source = nullptr;
            }
            break;
        }
    }
    switch (result)
    {
    case IDCANCEL:
    case IDNO:
        ::MessageBox(IDS_AFXBARRES_CANCEL, MB_ICONINFORMATION);
        break;
    case IDABORT:
        ::MessageBox(IDS_EDIT_ABORT, MB_ICONINFORMATION);
        break;
    }
    return result;
}


HICON GetToolBarIcon(UINT nID)
{
    HICON result = nullptr;
    // フレームウィンドウへのハンドルを取得
    HWND hWnd = ::m_hMainWnd;
    if (::IsWindow(hWnd))
    {
        // ツールバーもしくは印刷プレビュー用ツールバーの取得
        UINT nValue = AFX_IDW_TOOLBAR;
        if (::IsMenu(::m_hMenu))
        {
            nValue = AFX_IDW_PREVIEW_BAR;
        }
        hWnd = ::GetDlgItem(::m_hMainWnd, nValue);
        if (::IsWindow(hWnd))
        {
            // イメージリストを取得
            HIMAGELIST source = (HIMAGELIST)::SendMessage(hWnd, TB_GETIMAGELIST, 0, 0);
            if (nullptr != source && INVALID_HANDLE_VALUE != source)
            {
                // ツールバーのボタン情報取得用構造体の構築
                size_t index = sizeof(TBBUTTONINFO);
                LPTBBUTTONINFO value = (LPTBBUTTONINFO)::malloc(index);
                if (nullptr != value)
                {
                    ::SecureZeroMemory(value, index);
                    value->cbSize = (UINT)index;
                    value->dwMask = TBIF_IMAGE;

                    // コマンドIDを与えてイメージIDを取得
                    if (-1 != ::SendMessage(hWnd, TB_GETBUTTONINFO, nID, (LPARAM)value))
                    {
                        // イメージリストからアイコンを取得する
                        result = ::ImageList_GetIcon(source, value->iImage, ILD_NORMAL);
                    }

                    // ツールバーのボタン情報取得用構造体の解放
                    ::SecureZeroMemory(value, index);
                    ::free(value);
                    value = nullptr;
                }
            }
        }
    }
    return result;
}


int CommandToIndex(UINT nID)
{
    int result = IDC_STATIC;
    if (nullptr != ::m_pStatusBar)
    {
        HWND source = ::m_hMainWnd;
        if (::IsWindow(source))
        {
            source = ::GetDlgItem(source, AFX_IDW_STATUS_BAR);
            if (::IsWindow(source))
            {
                int value = (int)::SendMessage(source, SB_GETPARTS, 0, 0);
                for (int index = 0; IDC_STATIC == result && index < value; index++)
                {
                    if (::m_pStatusBar[index] == nID)
                    {
                        result = index;
                    }
                };
            }
        }
    }
    return result;
}


BOOL SaveModified()
{
    BOOL result = TRUE;
    if (::m_bDirty)
    {
        // ファイルタイトルもしくは「無題」の取得
        LPCTSTR value = nullptr;
        if (nullptr != ::m_pOpenFileName &&
            nullptr != ::m_pOpenFileName->lpstrFileTitle &&
            NULLSTR != *::m_pOpenFileName->lpstrFileTitle)
        {
            // ファイルタイトルがある場合、複製を取得
            value = _tcsdup(::m_pOpenFileName->lpstrFileTitle);
        }
        else
        {
            // ファイルタイトルが無い場合、「IDR_MAINFRAME」リソースから取得
            LPTSTR source = (LPTSTR)::m_szWindowClass;
            if (nullptr != source)
            {
                if (0 < ::LoadString(::m_hInst, IDR_MAINFRAME, source, 1 + MAX_PATH))
                {
                    // 1 == pos を取得。「タイトルなし」が指定されている場合がある
                    value = ::AfxExtractSubString(source, 1);
                    if (nullptr != value && NULLSTR == *value)
                    {
                        // IDR_MAINFRAME の 1 == pos が、空文字の場合、取得した文字を解放する
                        ::free((void*)value);
                        value = nullptr;
                    }
                }
                // リソースが取得できないか、空文字の場合、メインリソースから「無題」を複製して取得
                if (nullptr == value && 0 < ::LoadString(::m_hInst, AFX_IDS_UNTITLED, source, 1 + MAX_PATH))
                {
                    value = ::_tcsdup((LPCTSTR)source);
                }
            }
        }
        UINT index = IDCANCEL;
        if (nullptr != value)
        {
            // 文字列を書式化する
            LPCTSTR source = ::AfxFormatString(AFX_IDP_ASK_TO_SAVE, value);
            if (nullptr != source)
            {
                // メッセージボックスを起動
                index = ::MessageBox(source, MB_YESNOCANCEL | MB_ICONQUESTION | MB_DEFBUTTON3);
                // 書式化されたメッセージボックス用タイトルを解放
                ::free((void*)source);
                source = nullptr;
            }
            // タイトルもしくは「無題」を解放
            ::free((void*)value);
            value = nullptr;
        }
        result = FALSE;
        // メインフレームへのハンドルを取得しておく
        HWND source = ::m_hMainWnd;
        switch (index)
        {
        case IDYES:
            if (::IsWindow(source))
            {
                // ドキュメントを「保存」へ遷移
                ::SendMessage(source, WM_COMMAND, MAKEWPARAM(ID_FILE_SAVE, 0), 0);
            }
            result = FALSE == ::m_bDirty;
            break;
        case IDNO:
            result = TRUE;
            break;
        }
    }
    return result;
}


int GetSubItemCount()
{
    int result = 0;
    HWND value = ::m_hMainWnd;
    if (::IsWindow(value))
    {
        value = ::GetDlgItem(value, AFX_IDW_PANE_FIRST);
        if (::IsWindow(value))
        {
            value = (HWND)::SendMessage(value, LVM_GETHEADER, 0, 0);
            if (::IsWindow(value))
            {
                result = (int)::SendMessage(value, HDM_GETITEMCOUNT, 0, 0);
            }
        }
    }
    return result;
}


LPCTSTR DDX_ListViewItemText(HWND hWnd, int iSubItem, LPCTSTR source)
{
    LPCTSTR result = nullptr;
    if (::IsWindow(hWnd))
    {
        UINT index = sizeof(LVITEM);
        LPLVITEM pLvItem = (LPLVITEM)::malloc(index);
        if (nullptr != pLvItem)
        {
            ::SecureZeroMemory(pLvItem, index);
            pLvItem->mask       = LVIF_TEXT;
            pLvItem->iItem      = ::m_iItem;
            pLvItem->iSubItem   = iSubItem;

            if (nullptr == source)
            {
                for (int index = 31; nullptr == result && (UINT)index <= (UINT)INT_MAX; index <<= 1, index++)
                {
                    pLvItem->cchTextMax = index;
                    size_t value = sizeof(TCHAR) * index;
                    pLvItem->pszText = (LPTSTR)::malloc(value);
                    if (nullptr != pLvItem->pszText)
                    {
                        ::SecureZeroMemory(pLvItem->pszText, value);

                        if (FALSE != ::SendMessage(hWnd, LVM_GETITEM, 0, (LPARAM)pLvItem))
                        {
                            value = ::_tcslen((LPCTSTR)pLvItem->pszText);
                            if (0 < value && value < (size_t)(index - 3))
                            {
                                result = ::_tcsdup(pLvItem->pszText);
                            }
                        }

                        ::SecureZeroMemory(pLvItem->pszText, value);
                        ::free(pLvItem->pszText);
                        pLvItem->pszText = nullptr;
                    }
                };
            }
            else
            {
                pLvItem->mask       |= LVIF_IMAGE;
                pLvItem->pszText    = (LPTSTR)source;
                pLvItem->iImage     = 4;
                if (0 == iSubItem && (int)::SendMessage(hWnd, LVM_GETITEMCOUNT, 0, 0) <= ::m_iItem)
                {
                    pLvItem->mask   |= LVIF_PARAM;
                    pLvItem->lParam = ::m_iItem;
                    if (::m_iItem == (int)::SendMessage(hWnd, LVM_INSERTITEM, 0, (LPARAM)pLvItem))
                    {
                        result = source;
                    }
                }
                else
                {
                    if (FALSE != ::SendMessage(hWnd, LVM_SETITEM, 0, (LPARAM)pLvItem))
                    {
                        result = source;
                    }
                }
            }
            ::SecureZeroMemory(pLvItem, index);
            ::free(pLvItem);
            pLvItem = nullptr;
        }
    }
    return result;
}


void SetRowCol(HWND source)
{
    size_t index = sizeof(LVITEM);
    LPLVITEM value = (LPLVITEM)::malloc(index);
    if (nullptr != value)
    {
        ::SecureZeroMemory(value, index);
        value->iItem      = ::m_iItem;
        value->iSubItem   = 0;
        value->mask       = LVIF_STATE;
        value->state      = LVIS_SELECTED | LVIS_FOCUSED;
        value->stateMask  = value->state;
        ::SendMessage(source, LVM_SETITEM, 0, (LPARAM)value);
        ::free(value);
        value = nullptr;
    }
    ::SendMessage(source, LVM_ENSUREVISIBLE, (WPARAM)::m_iItem, (LPARAM)TRUE);
    ::SetFocus(source);
}


BOOL FindItem(HWND source, int index)
{
    BOOL result = FALSE;
    if (::IsWindow(source))
    {
        LPCTSTR value = DDX_ListViewItemText(source, index);
        if (nullptr != value)
        {
            LPFINDREPLACE source = ::m_pFindReplace;
            if (nullptr != source)
            {
                switch (source->Flags & FR_WHOLEWORD)
                {
                case FR_WHOLEWORD:
                    // 単語単位で探す（部分一致）
                    result = nullptr != ::_tcsstr(value, source->lpstrFindWhat);
                    break;
                default:
                    // 完全一致
                    result = 0 == ::_tcscmp(value, source->lpstrFindWhat);
                }
            }
            ::free((void*)value);
            value = nullptr;
        }
    }
    return result;
}


BSTR	str2BSTR(LPCTSTR	source)
{
    BSTR	result = nullptr;
    if (nullptr != source)
    {
#ifdef	_UNICODE
        result = ::SysAllocString(source);
#else
        size_t value = ::_tcsclen(source);
        int index = ::MultiByteToWideChar(CP_ACP, 0, source, (int)value, nullptr, 0);
        result = ::SysAllocStringLen(nullptr, index);
        if (nullptr != result)
        {
            ::MultiByteToWideChar(CP_ACP, 0, source, (int)value, result, index);
        }
#endif
    }
    return	result;
}


LPCTSTR	BSTR2str(BSTR	source)
{
    LPTSTR	result = nullptr;
    if (nullptr != source)
    {
#ifdef	_UNICODE
        result = ::_tcsdup((LPCWSTR)source);
#else
        int	index = ::WideCharToMultiByte(CP_ACP, 0, (OLECHAR*)source, -1, nullptr, 0, nullptr, nullptr);
        size_t value = (size_t)(index + 1);
        value *= sizeof(TCHAR);
        result = (LPTSTR)::malloc(value);
        if (nullptr != result)
        {
            ::WideCharToMultiByte(CP_ACP, 0, (OLECHAR*)source, -1, result, index, nullptr, nullptr);
        }
#endif
    }
    else
    {
        result = ::_tcsdup(_T(""));
    }
    return	(LPCTSTR)result;
}


HRESULT	create_instance(IDispatch** ppDispatch, LPCTSTR	source)
{
    HRESULT	result = E_FAIL;
    if (nullptr != source &&
        NULLSTR != *source &&
        nullptr != ppDispatch)
    {
        CLSID	clsid;
        BSTR	index = ::str2BSTR(source);
        if (nullptr != index)
        {
            result = ::CLSIDFromProgID(index, &clsid);
            if (FAILED(result))
            {
                result = ::CLSIDFromString(index, &clsid);
            }
            ::SysFreeString(index);
            index = nullptr;
        }
        if (SUCCEEDED(result))
        {
            result = ::CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)ppDispatch);
        }
    }
    return result;
}


HRESULT	com_invoke(IDispatch _In_* source, LPCTSTR _In_ pCommand,
    BOOL _In_ bPut, VARIANT _Out_opt_* pResult,
    VARIANT _In_opt_* pParam, UINT _In_ uParamCount)
{
    HRESULT result = E_FAIL;
    if (nullptr != source &&
        nullptr != pCommand &&
        NULLSTR != *pCommand)
    {
        DISPID	dispID = 0;
        BSTR	bstrValue = ::str2BSTR(pCommand);
        if (nullptr != bstrValue)
        {
            result = source->GetIDsOfNames(IID_NULL, &bstrValue, 1, LOCALE_USER_DEFAULT, (DISPID*)&dispID);
            ::SysFreeString(bstrValue);
            bstrValue = nullptr;
        }
        if (SUCCEEDED(result))
        {
            result = E_FAIL;
            USHORT uFlag = DISPATCH_PROPERTYGET;
            if (bPut)
            {
                uFlag = DISPATCH_PROPERTYPUT;
            }
            DISPPARAMS dispParams = { nullptr,	nullptr, 0,	0 };
            if (nullptr != pParam)
            {
                dispParams.cArgs = uParamCount;
                dispParams.rgvarg = pParam;
            }
            VARIANT objResult = {0}, * lResult = &objResult;
            if (nullptr != pResult)
            {
                lResult = pResult;
            }
            ::VariantInit(lResult);
            size_t value = sizeof(EXCEPINFO);
            LPEXCEPINFO	pExcepInfo = (LPEXCEPINFO)::malloc(value);
            if (nullptr != pExcepInfo)
            {
                ::SecureZeroMemory(pExcepInfo, value);

                UINT puArgErr = 0;
                result = source->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, uFlag, &dispParams, lResult, pExcepInfo, &puArgErr);

                ::free(pExcepInfo);
                pExcepInfo = nullptr;
            }
        }
    }
    return	result;
}


LPCTSTR GetDefaultConnect()
{
    LPCTSTR result = nullptr;
    if (nullptr != ::m_pOpenFileName &&
        nullptr != ::m_pOpenFileName->lpstrFile &&
        NULLSTR != *::m_pOpenFileName->lpstrFile)
    {
        int index = 1 + MAX_PATH;
        UINT value = IDS_ODBC_CONNECTION_X64;
#ifdef _M_IX86
        value = IDS_ODBC_CONNECTION_X86;
#endif
        if (0 < ::LoadString(::m_hInst, value, (LPTSTR)::m_szWindowClass, index))
        {
            ::_tcscat_s((LPTSTR)::m_szWindowClass, (size_t)index, ::m_pOpenFileName->lpstrFile);
            result = ::_tcsdup(::m_szWindowClass);
        }
    }
    return result;
}


LPCTSTR GetDefaultSQL()
{
    LPCTSTR result = nullptr;
    if (0 < ::LoadString(::m_hInst, IDS_SQL_QUERY, (LPTSTR)::m_szWindowClass, 1 + MAX_PATH))
    {
        result = ::_tcsdup(::m_szWindowClass);
    }
    return result;
}


BOOL    OpenDB()
{
    HRESULT result = ::create_instance(&::m_adoRS, _T("ADODB.Recordset"));
    if (SUCCEEDED(result))
    {
        IDispatch* adoCN = nullptr;
        result = ::create_instance(&adoCN, _T("ADODB.Connection"));
        if (SUCCEEDED(result))
        {
            VARIANT vResult, param;
            ::VariantInit(&vResult);
            ::VariantInit(&param);

            ::VariantClear(&param);
            param.vt = VT_I4;
            param.lVal = 3;//adUseClient
            result = ::com_invoke(adoCN, _T("CursorLocation"), TRUE, nullptr, &param);
            if (SUCCEEDED(result))
            {
                result = E_FAIL;
                LPCTSTR value = ::GetDefaultConnect();
                if (nullptr != value)
                {
                    ::VariantClear(&param);
                    param.vt = VT_BSTR;
                    param.bstrVal = ::str2BSTR(value);
                    if (nullptr != param.bstrVal)
                    {
                        result = ::com_invoke(adoCN, _T("ConnectionString"), TRUE, nullptr, &param);
                        ::SysFreeString(param.bstrVal);
                        param.bstrVal = nullptr;
                    }
                    ::free((void*)value);
                    value = nullptr;
                }
            }
            if (SUCCEEDED(result))
            {
                result = ::com_invoke(adoCN, _T("Open"));
            }
            if (SUCCEEDED(result))
            {
                ::VariantClear(&param);
                param.vt = VT_I4;
                param.lVal = 3;//adUseClient
                result = ::com_invoke(::m_adoRS, _T("CursorLocation"), TRUE, nullptr, &param);
                if (SUCCEEDED(result))
                {
                    ::VariantClear(&param);
                    param.vt = VT_I4;
                    param.lVal = 3;//adOpenStatic
                    result = ::com_invoke(::m_adoRS, _T("CursorType"), TRUE, nullptr, &param);
                }
                if (SUCCEEDED(result))
                {
                    ::VariantClear(&param);
                    param.vt = VT_I4;
                    param.lVal = 3;//adLockOptimistic
                    result = ::com_invoke(::m_adoRS, _T("LockType"), TRUE, nullptr, &param);
                }
                if (SUCCEEDED(result))
                {
                    result = E_FAIL;
                    LPCTSTR value = ::GetDefaultSQL();
                    if (nullptr != value)
                    {
                        ::VariantClear(&param);
                        param.vt = VT_BSTR;
                        param.bstrVal = ::str2BSTR(value);
                        if (nullptr != param.bstrVal)
                        {
                            result = ::com_invoke(::m_adoRS, _T("Source"), TRUE, nullptr, &param);
                            ::SysFreeString(param.bstrVal);
                            param.bstrVal = nullptr;
                        }
                        ::free((void*)value);
                        value = nullptr;
                    }
                }
                if (SUCCEEDED(result))
                {
                    ::VariantClear(&param);
                    param.vt = VT_DISPATCH | VT_BYREF;
                    param.ppdispVal = &adoCN;
                    result = ::com_invoke(::m_adoRS, _T("ActiveConnection"), TRUE, nullptr, &param);
                }
                if (SUCCEEDED(result))
                {
                    result = ::com_invoke(::m_adoRS, _T("Open"));
                }
                if (SUCCEEDED(result))
                {
                    IDispatch* dummy = nullptr;
                    ::VariantClear(&param);
                    param.vt = VT_DISPATCH | VT_BYREF;
                    param.ppdispVal = &dummy;
                    result = ::com_invoke(::m_adoRS, _T("ActiveConnection"), TRUE, nullptr, &param);
                }
                ::com_invoke(adoCN, _T("Close"));
                adoCN->Release();
                adoCN = nullptr;
            }
        }
    }
    return SUCCEEDED(result);
}


void    CloseDB()
{
    if (nullptr != ::m_adoRS)
    {
        if (::IsOpen())
        {
            ::com_invoke(::m_adoRS, _T("Close"));
        }
        ::m_adoRS->Release();
        ::m_adoRS = nullptr;
    }
}


BOOL    IsOpen()
{
    BOOL result = FALSE;
    if (nullptr != ::m_adoRS)
    {
        VARIANT objResult;
        if (SUCCEEDED(::com_invoke(::m_adoRS, _T("State"), FALSE, &objResult)))
        {
            result = 0 != objResult.lVal;
        }
    }
    return result;
}


BOOL    IsEOF()
{
    BOOL result = FALSE;
    if (nullptr != ::m_adoRS)
    {
        VARIANT objResult;
        if (SUCCEEDED(::com_invoke(::m_adoRS, _T("EOF"), FALSE, &objResult)))
        {
            result = 0 != objResult.lVal;
        }
    }
    return result;
}


void    MoveNext()
{
    if (nullptr != ::m_adoRS)
    {
        ::com_invoke(::m_adoRS, _T("MoveNext"));
    }
}


long    GetFieldCount()
{
    long result = 0;
    if (::IsOpen())
    {
        VARIANT vResult;
        if (SUCCEEDED(::com_invoke(::m_adoRS, _T("Fields"), FALSE, &vResult)))
        {
            ::m_adoFDS = vResult.pdispVal;
            if (SUCCEEDED(::com_invoke(::m_adoFDS, _T("Count"), FALSE, &vResult)))
            {
                result = vResult.lVal;
            }
        }
    }
    return result;
}


LPCTSTR GetFieldName(long index)
{
    LPCTSTR result = nullptr;
    if (::IsOpen() && nullptr != ::m_adoFDS)
    {
        VARIANT vResult, param;
        ::VariantInit(&param);
        param.vt = VT_I4;
        param.lVal = index;
        if (SUCCEEDED(::com_invoke(::m_adoFDS, _T("Item"), FALSE, &vResult, &param)) &&
            SUCCEEDED(::com_invoke(vResult.pdispVal, _T("Name"), FALSE, &vResult)))
        {
            switch (vResult.vt)
            {
            case VT_BSTR:
                result = ::BSTR2str(vResult.bstrVal);
                ::SysFreeString(vResult.bstrVal);
                break;
            }
        }
    }
    return result;
}


LPCTSTR GetFieldValue(long index)
{
    LPCTSTR result = nullptr;
    if (::IsOpen() && nullptr != ::m_adoFDS)
    {
        VARIANT vResult;
        VARIANT param;
        ::VariantInit(&param);
        param.vt = VT_I4;
        param.lVal = index;
        if (SUCCEEDED(::com_invoke(::m_adoFDS, _T("Item"), FALSE, &vResult, &param)) &&
            SUCCEEDED(::com_invoke(vResult.pdispVal, _T("Value"), FALSE, &vResult)))
        {
            switch (vResult.vt)
            {
            case VT_BSTR:
                result = ::BSTR2str(vResult.bstrVal);
                ::SysFreeString(vResult.bstrVal);
                break;
            }
        }
    }
    return result;
}


LPCTSTR SetPrinter(HGLOBAL hDevMode, HWND hWnd)
{
    BOOL result = FALSE;
    LPCTSTR pMessage = _T("DEVMODE ハンドルが指定されていません");
    if (nullptr != hDevMode && INVALID_HANDLE_VALUE != hDevMode)
    {
        pMessage = _T("DEVMODE ポインタの取得に失敗しました");
        LPDEVMODE pDevMode = (LPDEVMODE)::GlobalLock(hDevMode);
        if (nullptr != pDevMode)
        {
            pMessage = _T("デバイス名の取得に失敗しました");
            //デバイス名の取得
            int index = 1 + (int)::_tcslen(pDevMode->dmDeviceName);
            UINT value = sizeof(TCHAR) * index;
            LPTSTR szPrinterName = (LPTSTR)::malloc(value);
            if (nullptr != szPrinterName)
            {
                ::_tcscpy_s(szPrinterName, index, 
                    (LPCTSTR)(pDevMode->dmDeviceName));
                //PRINTER_ALL_ACCESS権でプリンタオープン
                pMessage = _T("プリンタオープン属性情報設定用バッファの構築に失敗しました");
                PRINTER_DEFAULTS* pPrinterDefault = (PRINTER_DEFAULTS*)
                    ::malloc(sizeof(PRINTER_DEFAULTS));
                if (nullptr != pPrinterDefault)
                {
                    ::SecureZeroMemory(pPrinterDefault, sizeof(PRINTER_DEFAULTS));

                    pMessage = _T("プリンタのオープンに失敗しました");
                    HANDLE hPrinter = nullptr;
                    pPrinterDefault->DesiredAccess = PRINTER_ALL_ACCESS;
                    if (::OpenPrinter(szPrinterName, &hPrinter, pPrinterDefault))
                    {
                        //１回目はわざとに失敗して必要バイト数を取得
                        DWORD dwNeeded = 0;
                        pMessage =
                            _T("プリンタ情報取得用バッファの大きさの取得に失敗しました");
                        if (FALSE == ::GetPrinter(hPrinter, 2, 0, 0, &dwNeeded))
                        {
                            pMessage = _T("プリンタ情報取得用グローバルメモリーの確保に失敗しました");
                            DWORD dwSize = dwNeeded;
                            HGLOBAL hMem = ::GlobalAlloc(GHND, dwSize);
                            if (nullptr != hMem)
                            {
                                pMessage = _T("プリンタ情報取得用ローカルメモリーの取得に失敗しました");
                                LPPRINTER_INFO_2 pPrinterInfo2 = (LPPRINTER_INFO_2)::GlobalLock(hMem);
                                if (nullptr != pPrinterInfo2)
                                {
                                    //PRINTER_INFO_2構造体にプリンタ情報を取得
                                    pMessage = _T("プリンタ情報の取得に失敗しました");
                                    if (0 != ::GetPrinter(hPrinter, 2, (LPBYTE)pPrinterInfo2, dwSize, &dwNeeded))
                                    {
                                        //現在のプリンタ情報のうちユーザー入力されたものに変更
                                        pPrinterInfo2->pDevMode->dmFields
                                            = DM_ORIENTATION
                                            | DM_PAPERSIZE
                                            | DM_PAPERLENGTH
                                            | DM_PAPERWIDTH
                                            | DM_DEFAULTSOURCE
                                            | DM_PRINTQUALITY
                                            | DM_COLOR
                                            | DM_DUPLEX
                                            ;
                                        pPrinterInfo2->pDevMode->dmOrientation      = pDevMode->dmOrientation;
                                        pPrinterInfo2->pDevMode->dmPaperSize        = pDevMode->dmPaperSize;
                                        pPrinterInfo2->pDevMode->dmPaperLength      = pDevMode->dmPaperLength;
                                        pPrinterInfo2->pDevMode->dmPaperWidth       = pDevMode->dmPaperWidth;
                                        pPrinterInfo2->pDevMode->dmDefaultSource    = pDevMode->dmDefaultSource;
                                        pPrinterInfo2->pDevMode->dmPrintQuality     = pDevMode->dmPrintQuality;
                                        pPrinterInfo2->pDevMode->dmColor            = pDevMode->dmColor;
                                        // ダイアログが指定された場合
                                        // コントロールへ情報を設定
                                        if (::IsWindow(hWnd))
                                        {
                                            HWND hCtrl = ::GetDlgItem(hWnd, AFX_IDC_PRINT_DOCNAME);
                                            if (::IsWindow(hCtrl))
                                            {
                                                ::SetWindowText(hCtrl, ::m_pOpenFileName->lpstrFileTitle);
                                            }
                                            hCtrl = ::GetDlgItem(hWnd, AFX_IDC_PRINT_PRINTERNAME);
                                            if (::IsWindow(hCtrl))
                                            {
                                                ::SetWindowText(hCtrl, pDevMode->dmDeviceName);
                                            }
                                            hCtrl = ::GetDlgItem(hWnd, AFX_IDC_PRINT_PORTNAME);
                                            if (::IsWindow(hCtrl))
                                            {
                                                ::SetWindowText(hCtrl, pPrinterInfo2->pPortName);
                                            }
                                        }
                                        //プリンタ情報のセット
                                        pMessage = _T("プリンタ情報の設定に失敗しました");
                                        if (0 != ::SetPrinter(hPrinter, 2, (LPBYTE)pPrinterInfo2, 0))
                                        {
                                            result = TRUE;
                                        }
                                    }
                                    ::GlobalUnlock(hMem);
                                }
                                ::GlobalFree(hMem);
                            }
                        }
                        // 印刷中ダイアログが指定された場合、プリンタハンドルを閉じずに保持
                        if (::IsWindow(hWnd))
                        {
                            ::m_hPrinter = hPrinter;
                        }
                        else
                        {
                            ::ClosePrinter(hPrinter);
                        }
                    }
                    ::SecureZeroMemory(pPrinterDefault,
                        sizeof(PRINTER_DEFAULTS));
                    ::free(pPrinterDefault);
                    pPrinterDefault = nullptr;
                }
                ::SecureZeroMemory(szPrinterName, value);
                ::free(szPrinterName);
                szPrinterName = nullptr;
            }
            ::GlobalFree(hDevMode);
            ::GlobalUnlock(hDevMode);
            hDevMode = nullptr;
        }
    }
    return result ? nullptr: pMessage;
}

BOOL OnPrint(HDC hDC, int nPage)
{
    BOOL result = FALSE;
    if (nullptr != hDC &&
        INVALID_HANDLE_VALUE != hDC &&
        ::IsWindow(::m_hMainWnd))
    {
        // リストビューを取得
        HWND hCtrl = ::GetDlgItem(::m_hMainWnd, AFX_IDW_PANE_FIRST);
        if (::IsWindow(hCtrl))
        {
            // リストビューのアイテム個数とサブアイテムの個数を取得
            int iMaxItems = (int)::SendMessage(hCtrl, LVM_GETITEMCOUNT, 0, 0),
                iMaxSubItems = ::GetSubItemCount();

            // 右上マージン設定
            POINT pos = { ::m_nPrintMarginLeft,::m_nPrintMarginTop };

            // タイトル印刷

            // アプリケーションタイトルの印刷
            LPCTSTR value = ::m_szTitle;
            if (nullptr != value)
            {
                ::TextOut(hDC, pos.x, pos.y, value, (int)::_tcslen(value));
            }
            pos.x += ::m_nPrintFieldPitch;
            pos.x += ::m_nPrintFieldPitch;
            pos.x += ::m_nPrintFieldPitch;
            pos.x += ::m_nPrintFieldPitch;

            // ドキュメント名の印刷
            value = ::m_pOpenFileName->lpstrFileTitle;
            if (nullptr != value)
            {
                ::TextOut(hDC, pos.x, pos.y, value, (int)::_tcslen(value));
            }
            
            // 左マージン設定
            pos.x = ::m_nPrintMarginLeft;
            // 次行へ移動
            pos.y += ::m_nPrintRowPitch;

            // ヘッダ印刷

            // ヘッダ列情報取得構造体の構築
            int index = sizeof(LVCOLUMN);
            LPLVCOLUMN source = (LPLVCOLUMN)::malloc(index);
            if (nullptr != source)
            {
                ::SecureZeroMemory(source, index);

                source->mask         = LVCF_TEXT;
                source->cchTextMax   = 1 + MAX_PATH;
                source->pszText      = (LPTSTR)::m_szWindowClass;

                // ヘッダ情報
                result = TRUE;
                for (int iSubItem = 0; result && iSubItem < iMaxSubItems; iSubItem++)
                {
                    result = FALSE;
                    // サブアイテム列の情報を取得
                    if (FALSE != ::SendMessage(hCtrl, LVM_GETCOLUMN, iSubItem, (LPARAM)source))
                    {
                        // ヘッダ列を出力
                        value = source->pszText;
                        if (nullptr != value)
                        {
                            ::TextOut(hDC, pos.x, pos.y, value, (int)::_tcslen(value));
                        }

                        // 処理成功
                        result = TRUE;

                        // 次のサブアイテム列へ移動
                        pos.x += ::m_nPrintFieldPitch;
                    }
                };
                // 次行へ移動
                pos.y += ::m_nPrintRowPitch;

                // ヘッダ列情報取得構造体の解放
                ::SecureZeroMemory(source, index);
                ::free(source);
                source = nullptr;
            }

            // アイテム印刷

            // 現在のページ番号から印刷開始行を計算
            ::m_iItem = nPage;
            ::m_iItem--;
            ::m_iItem *= ::m_nPrintLineByPage;
            result = TRUE;
            for (int index = 0; result && index < ::m_nPrintLineByPage; index++)
            {
                pos.x = ::m_nPrintMarginLeft;//左マージン
                if (::m_iItem < iMaxItems)
                {
                    for (int iSubItem = 0; result && iSubItem < iMaxSubItems; iSubItem++)
                    {
                        // 現在のサブアイテム位置の文字列を取得
                        LPCTSTR value = ::DDX_ListViewItemText(hCtrl, iSubItem);
                        if (nullptr != value)
                        {
                            // サブアイテムを出力
                            ::TextOut(hDC, pos.x, pos.y, value, (int)::_tcslen(value));

                            // 処理成功
                            result = TRUE;

                            // 現在のサブアイテム位置の文字列を解放
                            ::free((void*)value);
                            value = nullptr;
                        }
                        // 次のサブアイテム列へ移動
                        pos.x += ::m_nPrintFieldPitch;
                    };
                    ::m_iItem++;
                }
                // 次行へ移動
                pos.y += ::m_nPrintRowPitch;
            };
            if (result)
            {
                result = FALSE;

                // フッタ印刷

                // フッタ作成
                int index = 1 + MAX_PATH;
                // 書式文字列
                LPCTSTR source = nullptr;
                LPTSTR value = (LPTSTR)::m_szWindowClass;
                if (nullptr != value)
                {
                    // 書式文字列取得
                    if (0 < ::LoadString(::m_hInst, AFX_IDS_PRINTPAGENUM, value, index))
                    {
                        source = ::_tcsdup((LPCTSTR)value);
                    }
                }
                // 書式文字列取得判定
                if (nullptr != source)
                {
                    // 戻り値用バッファ作成
                    LPTSTR value = (LPTSTR)::m_szWindowClass;
                    if (nullptr != value)
                    {
                        // 書式文字列に従ってページ番号文字列を作成
                        ::_stprintf_s(value, (size_t)index, source, (UINT)nPage);
                        // フッタ出力
                        pos.x = ::m_nPrintMarginLeft;
                        ::TextOut(hDC, pos.x, pos.y, value, (int)::_tcslen(value));
                        // 戻り値
                        result = TRUE;
                    }
                    // 書式文字列解放
                    ::free((void*)source);
                    source = nullptr;
                }
            }
        }
    }
    return result;
}
