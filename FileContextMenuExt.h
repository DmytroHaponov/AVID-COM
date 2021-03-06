/****************************** Module Header ******************************\
Module Name:  FileContextMenuExt.h
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
to the context menu when you right-click selected files in the Windows Explorer. 
Clicking the menu item brings up a message box that displays the full path 
of selected files.

\***************************************************************************/

#pragma once

#include <windows.h>
#include <shlobj.h>     // For IShellExtInit and IContextMenu
#include <string>
#include <vector>
#include <thread>
#include <mutex>
#include <set>


class FileContextMenuExt : public IShellExtInit, public IContextMenu
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IShellExtInit
    IFACEMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, 
                                                               HKEY hKeyProgID);

    // IContextMenu
    IFACEMETHODIMP QueryContextMenu(HMENU hMenu, UINT indexMenu, 
                                    UINT idCmdFirst, UINT idCmdLast,
                                    UINT uFlags);
    IFACEMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    IFACEMETHODIMP GetCommandString(UINT_PTR idCommand, UINT uFlags,
                                    UINT *pwReserved, LPSTR pszName,
                                    UINT cchMax);
    
    FileContextMenuExt(void);

protected:
    ~FileContextMenuExt(void);

private:
    // Reference count of component.
    long m_cRef;

//! Haponov: container for displaying files info
    std::set<std::wstring> sortedFiles;
//! Haponov: container for full paths of selected files,
//! is used provide this info to threads of void processSelectedFiles(ws_name)
    std::vector<std::wstring> filePaths;

    /*       not used anymore
//! Haponov: convert string to wstring
    std::wstring s2ws(const std::string& s);
     */
//! Haponov: get file creation time
    BOOL GetCreationTime(HANDLE hFile, LPTSTR lpszString, DWORD dwSize);

//! Haponov: calculate checksum
    DWORD getCheckSum(std::wstring path);

    // The method that handles the "display" verb.
    void OnVerbDisplayFileName(HWND hWnd);

    PWSTR m_pszMenuText;
    HANDLE m_hMenuBmp;
    PCSTR m_pszVerb;
    PCWSTR m_pwszVerb;
    PCSTR m_pszVerbCanonicalName;
    PCWSTR m_pwszVerbCanonicalName;
    PCSTR m_pszVerbHelpText;
    PCWSTR m_pwszVerbHelpText;

//! Haponov: mutex for concurrency control
    std::mutex mu;

//! Haponov: process file info:
    //1) get file { name, size, creation date, ala checksum },
    //2) sort file names with according info
    void processSelectedFiles(std::wstring ws_name);
};