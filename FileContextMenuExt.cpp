/****************************** Module Header ******************************\
Module Name:  FileContextMenuExt.cpp
Project:      CppShellExtContextMenuHandler

The code sample demonstrates creating a Shell context menu handler with C++.

A context menu handler is a shell extension handler that adds commands to an
existing context menu. Context menu handlers are associated with a particular
file class and are called any time a context menu is displayed for a member
of the class. While you can add items to a file class context menu with the
registry, the items will be the same for all members of the class. By
implementing and registering such a handler, you can dynamically add items to
an object's context menu, customized for the particular object.

The example context menu handler adds the menu item "Avid the Best"
to the context menu when you right-click any number of selected files in the 
Windows Explorer.
Clicking the menu item brings up a message box that displays the full path
of selected files.

\***************************************************************************/

#include "FileContextMenuExt.h"
#include "resource.h"
#include <strsafe.h>
#include <Shlwapi.h>

#include <cstddef>
#include <sstream>

#include <tchar.h>
#include <fstream>

#pragma comment(lib, "shlwapi.lib")


extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0  // The command's identifier offset

FileContextMenuExt::FileContextMenuExt(void) : m_cRef(1),
//! Haponov change names

m_pszMenuText(L"&Avid the Best"),
m_pszVerb("cppdisplay"),
m_pwszVerb(L"cppdisplay"),
m_pszVerbCanonicalName("CppDisplayFileName"),
m_pwszVerbCanonicalName(L"CppDisplayFileName"),
m_pszVerbHelpText("Avid the Best"),
m_pwszVerbHelpText(L"Avid the Best")
//! end of Haponov change names
{
    InterlockedIncrement(&g_cDllRef);

    // Load the bitmap for the menu item. 
    // If you want the menu item bitmap to be transparent, the color depth of 
    // the bitmap must not be greater than 8bpp.
    m_hMenuBmp = LoadImage(g_hInst, MAKEINTRESOURCE(IDB_OK),
        IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE | LR_LOADTRANSPARENT);
}

FileContextMenuExt::~FileContextMenuExt(void)
{
    if (m_hMenuBmp)
    {
        DeleteObject(m_hMenuBmp);
        m_hMenuBmp = NULL;
    }

    InterlockedDecrement(&g_cDllRef);
}

//! Haponov function
std::wstring FileContextMenuExt::s2ws(const std::string & s)
{
    int len;
    int slength = (int)s.length() + 1;
    len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, 0, 0);
    wchar_t* buf = new wchar_t[len];
    MultiByteToWideChar(CP_ACP, 0, s.c_str(), slength, buf, len);
    std::wstring r(buf);
    delete[] buf;
    return r;
}

//! Haponov function - modified other msdn code sample
BOOL FileContextMenuExt::GetCreationTime(HANDLE hFile, 
                                                LPTSTR lpszString, DWORD dwSize)
{
    FILETIME ftCreate, ftAccess, ftWrite;
    SYSTEMTIME stUTC, stLocal;
    DWORD dwRet;

    // Retrieve the file times for the file.
    if (!GetFileTime(hFile, &ftCreate, &ftAccess, &ftWrite))
        return FALSE;

    // Convert the creation time to local time.
    FileTimeToSystemTime(&ftCreate, &stUTC);
    SystemTimeToTzSpecificLocalTime(NULL, &stUTC, &stLocal);

    // Build a string showing the date and time.
    dwRet = StringCchPrintfW(lpszString, dwSize, TEXT("%02d/%02d/%d %02d:%02d"),
            stLocal.wMonth, stLocal.wDay, stLocal.wYear, stLocal.wHour, 
            stLocal.wMinute);

    if (S_OK == dwRet)
        return TRUE;
    else return FALSE;
}

//! Haponov function
DWORD FileContextMenuExt::getCheckSum(wchar_t * path)
{
    DWORD checksum = 0;
    char byte;
    std::ifstream in(path, std::ifstream::binary);
    if (!in)
        return 0;
    else
    {
        while (in)
        {
            in.get(byte);
            checksum += byte;
        }
    }
    return checksum;
}

void FileContextMenuExt::OnVerbDisplayFileName(HWND hWnd)
{
    //! Haponov changes start here:

    //! Haponov: create Message text from all strings of sortedFiles std::set
    std::string sum;
    std::set<std::string>::const_iterator i = sortedFiles.begin();
    while (i != sortedFiles.end())
    {
        sum += *i; if (sum.empty()) { sum = "empty";	break; }
        if (++i != sortedFiles.end()) sum += "\n\n";
    }
    //! Haponov: prepare message string to be sent to MessageBox
    std::wstring stemp = s2ws(sum);
    LPCWSTR msg = stemp.c_str();
    MessageBox(hWnd, msg, L"CppShellExtContextMenuHandler", MB_OK);
    //! end of Haponov changes
}

//! Haponov function
void FileContextMenuExt::processSelectedFiles(wchar_t * path)
{
    //! Haponov: all string formats must be converted to std::string 
    //! in order to conveniently use strings in STL containers
    //!------------------------------

    //! Haponov: convert full path of file to string
    std::wstring ws_name(path);
    std::string atLast(ws_name.begin(), ws_name.end()); 
    //! Haponov: get a string with filename only - WO full path
    std::size_t found = atLast.find_last_of("/\\");
    atLast = atLast.substr(found + 1);					

    //! Haponov: create file handle for further usage
    HANDLE hFile = CreateFile(ws_name.c_str(),
        GENERIC_READ,
        FILE_SHARE_READ,
        NULL,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        NULL);
    if (hFile == INVALID_HANDLE_VALUE)
    {
        atLast = "error opening file";
        return;
    }

    // Haponov: get a string with size of file
    DWORD dwFileSize = GetFileSize(hFile, NULL);
    std::ostringstream streamSize;
    streamSize << dwFileSize;
    std::string result_size = streamSize.str();
    // Haponov: put spaces into size - "размер в удобочитаемом виде"
    unsigned int curLength = result_size.length();
    if (curLength > 3)
    {
        curLength -= 3;
        for (int i=3; curLength < result_size.length(); i += 4, curLength -= 3)
        {
            if (curLength > result_size.length() || !curLength) break;
            result_size.insert(result_size.end() - i, ' ');
        }
    }

    // Haponov: get a string with creation time of file
    wchar_t temp_forCreationTime[MAX_PATH];
    if (!GetCreationTime(hFile, temp_forCreationTime, 
                                            ARRAYSIZE(temp_forCreationTime)))
        return;
    std::wstring ws_creationTime(temp_forCreationTime);
    std::string resultCreationTime(ws_creationTime.begin(), 
                                                        ws_creationTime.end());	

    // Haponov: get a string with ala checksum
    DWORD checksum = getCheckSum(path);
    std::ostringstream streamCheckSum;
    streamCheckSum << checksum;
    std::string result_checkSum = streamCheckSum.str();

    // Haponov: create a resulting string for displaying
    atLast += ";   size: ";   atLast += result_size;  
    atLast += " KB;   creation time: ";  	atLast += resultCreationTime; 
    atLast += "   checksum: ";   atLast += result_checkSum;

    // Haponov: protect shared std::set<std::string> sortedFiles
    mu.lock(); 
    sortedFiles.insert(atLast);
    mu.unlock();
    return;
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP FileContextMenuExt::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] =
    {
        QITABENT(FileContextMenuExt, IContextMenu),
        QITABENT(FileContextMenuExt, IShellExtInit),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FileContextMenuExt::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FileContextMenuExt::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion


#pragma region IShellExtInit

// Initialize the context menu handler.
IFACEMETHODIMP FileContextMenuExt::Initialize(
    LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    if (NULL == pDataObj)
    {
        return E_INVALIDARG;
    }

    HRESULT hr = E_FAIL;

    FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

    // The pDataObj pointer contains the objects being acted upon. In this 
    // example, we get an HDROP handle for enumerating the selected files and 
    // folders.
    if (SUCCEEDED(pDataObj->GetData(&fe, &stm)))
    {
        // Get an HDROP handle.
        HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
        if (hDrop != NULL)
        {
            //! start of Haponov CHANGES:
            //!****************************************************
            //! Haponov: number of threads depends on number of selected files 
            UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
            if (nFiles)
            {
                if (nFiles > 1)
                {
                    std::vector<std::thread> threads;

                    for (int i = 0; i < nFiles; ++i)
                    {
                        wchar_t temp_forName[MAX_PATH];
                        // Get the name of the file.
                        if (0 != DragQueryFile(hDrop, i, temp_forName 
                            /*such path is written to temp_forName*/,
                            ARRAYSIZE(temp_forName)))
                        {
                            threads.push_back(std::thread(&FileContextMenuExt::
                                    processSelectedFiles, this, temp_forName));
                        }
                    }
                    //! Haponov: use the main thread to do part of the work
                    wchar_t temp1_forName[MAX_PATH];
                        if (0 != DragQueryFile(hDrop, 0, temp1_forName,
                                                    ARRAYSIZE(temp1_forName)))
                        {
                            processSelectedFiles(temp1_forName);
                        }
                    
                    for (auto &t : threads)
                    {
                        t.join();
                    }
                }
                //! Haponov: use only main thread if a sinle file is selected
                else
                {
                    wchar_t temp_forName[MAX_PATH];
                    for (int i = nFiles - 1; i < nFiles; ++i)
                    {
                        if (0 != DragQueryFile(hDrop, i, temp_forName,
                                                       ARRAYSIZE(temp_forName)))
                            processSelectedFiles(temp_forName);
                    }
                }
            }
            if (sortedFiles.size()) hr = S_OK;
            //! end of Haponov changes 
            //!****************************************************
            GlobalUnlock(stm.hGlobal);
        }
        ReleaseStgMedium(&stm);
    }

    // If any value other than S_OK is returned from the method, the context 
    // menu item is not displayed.
    return hr;
}

#pragma endregion


#pragma region IContextMenu

//
//   FUNCTION: FileContextMenuExt::QueryContextMenu
//
//   PURPOSE: The Shell calls IContextMenu::QueryContextMenu to allow the 
//            context menu handler to add its menu items to the menu. It 
//            passes in the HMENU handle in the hmenu parameter. The 
//            indexMenu parameter is set to the index to be used for the 
//            first menu item that is to be added.
//
IFACEMETHODIMP FileContextMenuExt::QueryContextMenu(
    HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    // If uFlags include CMF_DEFAULTONLY then we should not do anything.
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

    // Use either InsertMenu or InsertMenuItem to add menu items.

    MENUITEMINFO mii = { sizeof(mii) };
    mii.fMask = MIIM_BITMAP | MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
    mii.wID = idCmdFirst + IDM_DISPLAY;
    mii.fType = MFT_STRING;
    mii.dwTypeData = m_pszMenuText;
    mii.fState = MFS_ENABLED;
    mii.hbmpItem = static_cast<HBITMAP>(m_hMenuBmp);
    if (!InsertMenuItem(hMenu, indexMenu, TRUE, &mii))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Add a separator.
    MENUITEMINFO sep = { sizeof(sep) };
    sep.fMask = MIIM_TYPE;
    sep.fType = MFT_SEPARATOR;
    if (!InsertMenuItem(hMenu, indexMenu + 1, TRUE, &sep))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
    // Set the code value to the offset of the largest command identifier 
    // that was assigned, plus one (1).
    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(IDM_DISPLAY + 1));
}


//
//   FUNCTION: FileContextMenuExt::InvokeCommand
//
//   PURPOSE: This method is called when a user clicks a menu item to tell 
//            the handler to run the associated command. The lpcmi parameter 
//            points to a structure that contains the needed information.
//
IFACEMETHODIMP FileContextMenuExt::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    BOOL fUnicode = FALSE;

    // Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
    // CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
    // the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
    // structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
    // structure. This struct is an extended version of CMINVOKECOMMANDINFO 
    // and has additional members that allow Unicode strings to be passed.
    if (pici->cbSize == sizeof(CMINVOKECOMMANDINFOEX))
    {
        if (pici->fMask & CMIC_MASK_UNICODE)
        {
            fUnicode = TRUE;
        }
    }

    // Determines whether the command is identified by its offset or verb.
    // There are two ways to identify commands:
    // 
    //   1) The command's verb string 
    //   2) The command's identifier offset
    // 
    // If the high-order word of lpcmi->lpVerb (for the ANSI case) or 
    // lpcmi->lpVerbW (for the Unicode case) is nonzero, lpVerb or lpVerbW 
    // holds a verb string. If the high-order word is zero, the command 
    // offset is in the low-order word of lpcmi->lpVerb.

    // For the ANSI case, if the high-order word is not zero, the command's 
    // verb string is in lpcmi->lpVerb. 
    if (!fUnicode && HIWORD(pici->lpVerb))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIA(pici->lpVerb, m_pszVerb) == 0)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // For the Unicode case, if the high-order word is not zero, the 
    // command's verb string is in lpcmi->lpVerbW. 
    else if (fUnicode && HIWORD(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW))
    {
        // Is the verb supported by this context menu extension?
        if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, m_pwszVerb) == 0)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    // If the command cannot be identified through the verb string, then 
    // check the identifier offset.
    else
    {
        // Is the command identifier offset supported by this context menu 
        // extension?
        if (LOWORD(pici->lpVerb) == IDM_DISPLAY)
        {
            OnVerbDisplayFileName(pici->hwnd);
        }
        else
        {
            // If the verb is not recognized by the context menu handler, it 
            // must return E_FAIL to allow it to be passed on to the other 
            // context menu handlers that might implement that verb.
            return E_FAIL;
        }
    }

    return S_OK;
}


//
//   FUNCTION: CFileContextMenuExt::GetCommandString
//
//   PURPOSE: If a user highlights one of the items added by a context menu 
//            handler, the handler's IContextMenu::GetCommandString method is 
//            called to request a Help text string that will be displayed on 
//            the Windows Explorer status bar. This method can also be called 
//            to request the verb string that is assigned to a command. 
//            Either ANSI or Unicode verb strings can be requested. This 
//            example only implements support for the Unicode values of 
//            uFlags, because only those have been used in Windows Explorer 
//            since Windows 2000.
//
IFACEMETHODIMP FileContextMenuExt::GetCommandString(UINT_PTR idCommand,
    UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    HRESULT hr = E_INVALIDARG;
    
    if (idCommand == IDM_DISPLAY)
    {
        switch (uFlags)
        {
        case GCS_HELPTEXTW:
            // Only useful for pre-Vista versions of Windows that have a 
            // Status bar.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax,
                m_pwszVerbHelpText);
            break;

        case GCS_VERBW:
            // GCS_VERBW is an optional feature that enables a caller to 
            // discover the canonical name for the verb passed in through 
            // idCommand.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax,
                m_pwszVerbCanonicalName);
            break;

        default:
            hr = S_OK;
        }
    }

    // If the command (idCommand) is not supported by this context menu 
    // extension handler, return E_INVALIDARG.

    return hr;
}

#pragma endregion