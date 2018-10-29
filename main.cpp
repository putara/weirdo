#define WINVER          0x0600
#define _WIN32_WINNT    0x0600
#define _WIN32_IE       0x0700
#define UNICODE
#define _UNICODE

#include <sdkddkver.h>
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <commoncontrols.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <wincodec.h>
#include <uxtheme.h>
#include <stddef.h>
#include <intrin.h>

#include <strsafe.h>

//#define USETRACE

#define LZ4_COMPRESS_MODE       1   // 1: LZ4, 2: LZ4 HC

#include "lib/lz4/lib/lz4.h"
#include "lib/lz4/lib/lz4hc.h"
#include "lib/match/match.h"

// https://blogs.msdn.microsoft.com/oldnewthing/20041025-00/?p=37483/
extern "C" IMAGE_DOS_HEADER     __ImageBase;
#define HINST_THISCOMPONENT     ((HINSTANCE)(&__ImageBase))

/* void Cls_OnSettingChange(HWND hwnd, UINT uiParam, LPCTSTR lpszSectionName) */
#define HANDLE_WM_SETTINGCHANGE(hwnd, wParam, lParam, fn) \
    ((fn)((hwnd), (UINT)(wParam), (LPCTSTR)(lParam)), 0L)
#define FORWARD_WM_SETTINGCHANGE(hwnd, uiParam, lpszSectionName, fn) \
    (void)(fn)((hwnd), WM_WININICHANGE, (WPARAM)(UINT)(uiParam), (LPARAM)(LPCTSTR)(lpszSectionName))

// HANDLE_WM_CONTEXTMENU is buggy
// https://blogs.msdn.microsoft.com/oldnewthing/20040921-00/?p=37813
/* void Cls_OnContextMenu(HWND hwnd, HWND hwndContext, int xPos, int yPos) */
#undef HANDLE_WM_CONTEXTMENU
#undef FORWARD_WM_CONTEXTMENU
#define HANDLE_WM_CONTEXTMENU(hwnd, wParam, lParam, fn) \
    ((fn)((hwnd), (HWND)(wParam), (int)GET_X_LPARAM(lParam), (int)GET_Y_LPARAM(lParam)), 0L)
#define FORWARD_WM_CONTEXTMENU(hwnd, hwndContext, xPos, yPos, fn) \
    (void)(fn)((hwnd), WM_CONTEXTMENU, (WPARAM)(HWND)(hwndContext), MAKELPARAM((int)(xPos), (int)(yPos)))

#define FAIL_HR(hr) if (FAILED((hr))) goto Fail

template <class T>
void IUnknown_SafeAssign(T*& dst, T* src) throw()
{
    if (src != NULL) {
        src->AddRef();
    }
    if (dst != NULL) {
        dst->Release();
    }
    dst = src;
}

template <class T>
void IUnknown_SafeRelease(T*& unknown) throw()
{
    if (unknown != NULL) {
        unknown->Release();
        unknown = NULL;
    }
}

void Trace(__in __format_string const char* format, ...)
{
    TCHAR buffer[512];
    va_list argPtr;
    va_start(argPtr, format);
#ifdef _UNICODE
    WCHAR wformat[512];
    MultiByteToWideChar(CP_ACP, 0, format, -1, wformat, _countof(wformat));
    StringCchVPrintfW(buffer, _countof(buffer), wformat, argPtr);
#else
    StringCchVPrintfA(buffer, _countof(buffer), format, argPtr);
#endif
    WriteConsole(GetStdHandle(STD_OUTPUT_HANDLE), buffer, lstrlen(buffer), NULL, NULL);
    va_end(argPtr);
}

#undef TRACE
#undef ASSERT

#ifdef USETRACE
#define TRACE(s, ...)       Trace(s, __VA_ARGS__)
#define ASSERT(e)           do if (!(e)) { TRACE("%s(%d): Assertion failed\n", __FILE__, __LINE__); if (::IsDebuggerPresent()) { ::DebugBreak(); } } while(0)
#else
#define TRACE(s, ...)
#define ASSERT(e)
#endif


inline UINT32 PadUpPtr(UINT32 x) throw()
{
    return -static_cast<int>(x) & (sizeof(LONGLONG) - 1);
}

inline UINT32 AlignPtr(UINT32 x) throw()
{
    return (x + sizeof(LONGLONG) - 1) & ~(sizeof(LONGLONG) - 1);
}

inline UINT32 PadUp16(UINT32 x) throw()
{
    return -static_cast<int>(x) & 15;
}

inline UINT32 Align16(UINT32 x) throw()
{
    return (x + 15) & ~15;
}

inline UINT32 GetStride(UINT width, UINT bpp) throw()
{
    return ((width * bpp + 31) & ~31) / 8;
}


inline void* operator new(size_t size) throw()
{
    return ::malloc(size);
}

inline void* operator new[](size_t size) throw()
{
    return ::malloc(size);
}

inline void operator delete(void* ptr) throw()
{
    ::free(ptr);
}

inline void operator delete[](void* ptr) throw()
{
    ::free(ptr);
}

template <typename T>
inline HRESULT SafeAlloc(size_t size, __deref_out_bcount(size) T*& ptr) throw()
{
    ptr = static_cast<T*>(::malloc(size));
    return ptr != NULL ? S_OK : E_OUTOFMEMORY;
}

template <typename T>
inline void SafeFree(__deallocate_opt(Mem) T*& ptr) throw()
{
    ::free(ptr);
    ptr = NULL;
}

template <typename T>
inline HRESULT SafeAllocAligned(size_t size, __deref_out_bcount(size) T*& ptr) throw()
{
    ptr = static_cast<T*>(::_aligned_malloc(size, 16));
    return ptr != NULL ? S_OK : E_OUTOFMEMORY;
}

template <typename T>
inline void SafeFreeAligned(__deallocate_opt(Mem) T*& ptr) throw()
{
    ::_aligned_free(ptr);
    ptr = NULL;
}

template <class T>
inline HRESULT SafeNew(__deref_out T*& ptr) throw()
{
    ptr = new T();
    return ptr != NULL ? S_OK : E_OUTOFMEMORY;
}

template <class T>
inline void SafeDelete(__deallocate_opt(Mem) T*& ptr) throw()
{
    delete ptr;
    ptr = NULL;
}

inline HRESULT SafeStrDupW(const wchar_t* src, wchar_t*& dst) throw()
{
    dst = ::_wcsdup(src);
    return dst != NULL ? S_OK : E_OUTOFMEMORY;
}


template <class T>
class ComPtr
{
private:
    mutable T* ptr;

public:
    ComPtr() throw()
        : ptr()
    {
    }
    ~ComPtr() throw()
    {
        this->Release();
    }
    ComPtr(const ComPtr<T>& src) throw()
        : ptr()
    {
        operator =(src);
    }
    T* Detach() throw()
    {
        T* ptr = this->ptr;
        this->ptr = NULL;
        return ptr;
    }
    void Release() throw()
    {
        IUnknown_SafeRelease(this->ptr);
    }
    ComPtr<T>& operator =(T* src) throw()
    {
        IUnknown_SafeAssign(this->ptr, src);
        return *this;
    }
    operator T*() const throw()
    {
        return this->ptr;
    }
    T* operator ->() const throw()
    {
        return this->ptr;
    }
    T** operator &() throw()
    {
        ASSERT(this->ptr == NULL);
        return &this->ptr;
    }
    void CopyTo(__deref_out T** outPtr) throw()
    {
        ASSERT(this->ptr != NULL);
        (*outPtr = this->ptr)->AddRef();
    }
    HRESULT CoCreateInstance(REFCLSID clsid, IUnknown* outer, DWORD context) throw()
    {
        ASSERT(this->ptr == NULL);
        return ::CoCreateInstance(clsid, outer, context, IID_PPV_ARGS(&this->ptr));
    }
    template <class Q>
    HRESULT QueryInterface(__deref_out Q** outPtr) throw()
    {
        return this->ptr->QueryInterface(IID_PPV_ARGS(outPtr));
    }
};


template <class T>
class TPtrArrayDestroyHelper
{
public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_Destroy(hdpa);
    }
};

template <class T>
class TPtrArrayAutoDeleteHelper
{
    static int CALLBACK sDeletePtrCB(void* p, void*)
    {
        delete static_cast<T*>(p);
        return 1;
    }

public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_EnumCallback(hdpa, sDeletePtrCB, NULL);
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_DestroyCallback(hdpa, sDeletePtrCB, NULL);
    }
};

template <class T>
class TPtrArrayAutoFreeHelper
{
    static int CALLBACK sDeletePtrCB(void* p, void*)
    {
        ::free(p);
        return 1;
    }

public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_EnumCallback(hdpa, sDeletePtrCB, NULL);
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_DestroyCallback(hdpa, sDeletePtrCB, NULL);
    }
};

template <class T>
class TPtrArrayAutoCoFreeHelper
{
    static int CALLBACK sDeletePtrCB(void* p, void*)
    {
        ::SHFree(p);
        return 1;
    }

public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_EnumCallback(hdpa, sDeletePtrCB, NULL);
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_DestroyCallback(hdpa, sDeletePtrCB, NULL);
    }
};

template <class T, int cGrow, typename TDestroy>
class CPtrArray
{
private:
    HDPA        m_hdpa;

    CPtrArray(const CPtrArray&) throw();
    CPtrArray& operator =(const CPtrArray&) throw();

public:
    CPtrArray()
        : m_hdpa()
    {
    }
    ~CPtrArray()
    {
        TDestroy::Destroy(m_hdpa);
    }

    int GetCount() const throw()
    {
        if (m_hdpa == NULL) {
            return 0;
        }
        return DPA_GetPtrCount(m_hdpa);
    }

    T* FastGetAt(int index) const throw()
    {
        return static_cast<T*>(DPA_FastGetPtr(m_hdpa, index));
    }

    T* GetAt(int index) const throw()
    {
        return static_cast<T*>(::DPA_GetPtr(m_hdpa, index));
    }

    T** GetData() const throw()
    {
        return reinterpret_cast<T**>(DPA_GetPtrPtr(m_hdpa));
    }

    bool Grow(int nNewSize) throw()
    {
        return (::DPA_Grow(m_hdpa, nNewSize) != FALSE);
    }

    bool SetAtGrow(int index, T* p) throw()
    {
        return (::DPA_SetPtr(m_hdpa, index, p) != FALSE);
    }

    bool Create() throw()
    {
        if (m_hdpa != NULL) {
            return true;
        }
        m_hdpa = ::DPA_Create(cGrow);
        return (m_hdpa != NULL);
    }

    bool Create(HANDLE hHeap) throw()
    {
        if (m_hdpa != NULL) {
            return true;
        }
        m_hdpa = ::DPA_CreateEx(cGrow, hHeap);
        return (m_hdpa != NULL);
    }

    int Add(T* p) throw()
    {
        return InsertAt(DA_LAST, p);
    }

    int InsertAt(int index, T* p) throw()
    {
        return ::DPA_InsertPtr(m_hdpa, index, p);
    }

    T* RemoveAt(int index) throw()
    {
        return static_cast<T*>(::DPA_DeletePtr(m_hdpa, index));
    }

    void RemoveAll()
    {
        TDestroy::DeleteAllPtrs(m_hdpa);
    }

    void Enumerate(__in PFNDAENUMCALLBACK pfnCB, __in_opt void* pData) const
    {
        ::DPA_EnumCallback(m_hdpa, pfnCB, pData);
    }

    int Search(__in_opt void* pFind, __in int iStart, __in PFNDACOMPARE pfnCompare, __in LPARAM lParam, __in UINT options)
    {
        return ::DPA_Search(m_hdpa, pFind, iStart, pfnCompare, lParam, options);
    }

    bool Sort(__in PFNDACOMPARE pfnCompare, __in LPARAM lParam)
    {
        return (::DPA_Sort(m_hdpa, pfnCompare, lParam) != FALSE);
    }

    T* operator [](int index) const throw()
    {
        return GetAt(index);
    }
};


static HRESULT StreamRead(IStream* stream, void* data, DWORD size)
{
    if (size == 0) {
        return S_OK;
    }
    DWORD read = 0;
    HRESULT hr = stream->Read(data, size, &read);
    if (SUCCEEDED(hr) && size != read) {
        hr = E_FAIL;
    }
    return hr;
}

static HRESULT StreamWrite(IStream* stream, const void* data, DWORD size)
{
    if (size == 0) {
        return S_OK;
    }
    DWORD written = 0;
    HRESULT hr = stream->Write(data, size, &written);
    if (SUCCEEDED(hr) && size != written) {
        hr = E_FAIL;
    }
    return hr;
}


static const int PADDING = 8;
static const int SIZE32 = 32;
static const int SIZE16 = 16;
static const int IMGWIDTH = PADDING + SIZE32 + PADDING + SIZE16 + PADDING;
static const int IMGHEIGHT = PADDING + SIZE32 + PADDING;

static const UINT32 SIGNATURE = 'ALCI';

struct HEADER
{
    UINT32 signature;
    UINT32 reserved;
    UINT32 groups;
    UINT32 items;
    UINT32 imageoff;
    UINT32 imagesize;
    UINT32 nameoff;
    UINT32 namesize;
    // GROUP group[groups];
    // ITEM item[items];
    // IMAGE image[items];
    // NAME name[groups + items];
};

struct ITEM
{
    UINT32 image;
    UINT32 name;
};

struct GROUP
{
    UINT32 name;
    UINT32 count;
    ITEM items[ANYSIZE_ARRAY];
};

struct IMAGE
{
    UINT32 compsize;
    UINT32 orgsize;
    BYTE data[ANYSIZE_ARRAY];
};


class CacheBase
{
protected:
    static UINT32 SizeofString(const wchar_t* str) throw()
    {
        return static_cast<UINT32>(::wcslen(str) + 1) * 2;
    }

    static UINT32 SizeofStringAligned(const wchar_t* str) throw()
    {
        return AlignPtr(SizeofString(str));
    }

    static UINT32 SizeofGroup(UINT32 groups, UINT32 items) throw()
    {
        return Align16(offsetof(GROUP, items) * groups + sizeof(ITEM) * items);
    }

    static HRESULT StreamSeekTo(IStream* stream, UINT32 offset) throw()
    {
        LARGE_INTEGER pos;
        pos.LowPart = offset;
        pos.HighPart = 0;
        ULARGE_INTEGER upos;
        HRESULT hr = stream->Seek(pos, STREAM_SEEK_SET, &upos);
        if (SUCCEEDED(hr) && (upos.LowPart != offset || upos.HighPart != 0)) {
            hr = E_FAIL;
        }
        return hr;
    }

    static HRESULT StreamPadSkip(IStream* stream, UINT32 size) throw()
    {
        LARGE_INTEGER pad;
        pad.LowPart = PadUpPtr(size);
        pad.HighPart = 0;
        return stream->Seek(pad, STREAM_SEEK_CUR, NULL);
    }

    static HRESULT StreamWriteStr(IStream* stream, const wchar_t* str) throw()
    {
        UINT32 cb = SizeofString(str);
        HRESULT hr = StreamWrite(stream, str, cb);
        if (SUCCEEDED(hr)) {
            hr = StreamPadWrite(stream, cb);
        }
        return hr;
    }

    static HRESULT StreamPadWrite(IStream* stream, UINT32 size) throw()
    {
        LONGLONG zero = 0;
        UINT32 pad = PadUpPtr(size);
        return StreamWrite(stream, &zero, pad);
    }

    static HRESULT StreamPadWrite16(IStream* stream, UINT32 size) throw()
    {
        LONGLONG zeroes[2] = { 0, 0 };
        UINT32 pad = PadUp16(size);
        return StreamWrite(stream, zeroes, pad);
    }
};

class Cache : CacheBase
{
private:
    HEADER header;
    ComPtr<IStream> stream;
    GROUP** groups;
    wchar_t* names;
    wchar_t empty;

public:
    Cache() throw()
        : groups()
        , names()
        , empty()
    {
    }

    ~Cache() throw()
    {
        SafeFree(this->groups);
        SafeFree(this->names);
    }

    HRESULT Init(IStream* strm) throw()
    {
        HRESULT hr;
        this->stream = strm;
        FAIL_HR(hr = this->ReadHeader());
        FAIL_HR(hr = this->ReadGroup());
        FAIL_HR(hr = this->ReadName());
Fail:
        return hr;
    }

    UINT32 GetGroupCount() const throw()
    {
        return this->header.groups;
    }

    const wchar_t* GetGroupName(UINT32 group) const throw()
    {
        return group < this->header.groups ? this->names + this->groups[group]->name : &empty;
    }

    UINT32 GetAllItemCount() const throw()
    {
        return this->header.items;
    }

    UINT32 GetItemCount(UINT32 group) const throw()
    {
        return group < this->header.groups ? this->groups[group]->count : 0;
    }

    const ITEM* GetItem(UINT32 group, UINT32 item) const throw()
    {
        if (group < this->header.groups && item < this->groups[group]->count) {
            return this->groups[group]->items + item;
        }
        return NULL;
    }

    const wchar_t* GetItemName(UINT32 group, UINT32 item) const throw()
    {
        const ITEM* p = this->GetItem(group, item);
        if (p != NULL) {
            return this->names + p->name;
        }
        return &empty;
    }

    HRESULT CopyItemImage(const ITEM* item, __out_bcount(size) void* image, UINT32 size) const throw()
    {
        if (item != NULL) {
            HRESULT hr;
            const UINT32 off = this->header.imageoff + item->image;
            IMAGE hdr;
            char* tmp = NULL;

            FAIL_HR(hr = this->StreamSeekTo(off));
            FAIL_HR(hr = this->StreamRead(&hdr, offsetof(IMAGE, data)));
            FAIL_HR(hr = size >= hdr.orgsize ? S_OK : E_FAIL);

            if (hdr.orgsize == hdr.compsize) {
                FAIL_HR(hr = this->StreamRead(image, hdr.orgsize));
            } else {
                int copied;
                FAIL_HR(hr = SafeAlloc(hdr.compsize, tmp));
                FAIL_HR(hr = this->StreamRead(tmp, hdr.compsize));
                copied = ::LZ4_decompress_safe(tmp, static_cast<char*>(image), static_cast<int>(hdr.compsize), hdr.orgsize);
                FAIL_HR(hr = static_cast<UINT32>(copied) == hdr.orgsize ? S_OK : E_FAIL);
            }

Fail:
            SafeFree(tmp);
            return hr;
        }
        return E_FAIL;
    }

private:
    HRESULT ReadHeader() throw()
    {
        HRESULT hr;
        FAIL_HR(hr = this->StreamRead(&this->header, sizeof(this->header)));
        FAIL_HR(hr = this->header.signature == SIGNATURE ? S_OK : E_FAIL);
Fail:
        return hr;
    }

    HRESULT ReadGroup() throw()
    {
        const UINT32 sizeGroup = this->SizeofGroup();
        const UINT32 sizePtr = this->header.groups * sizeof(GROUP*);
        GROUP* group;
        HRESULT hr;
        FAIL_HR(hr = SafeAlloc(sizePtr + sizeGroup, this->groups));

        group = reinterpret_cast<GROUP*>(reinterpret_cast<BYTE*>(this->groups) + sizePtr);
        FAIL_HR(hr = this->StreamRead(group, sizeGroup));

        const UINT32 cGroups = this->header.groups;
        for (UINT32 i = 0; i < cGroups; i++) {
            this->groups[i] = group;
            UINT32 c = group->count;
            ITEM* item = group->items;
            for (; c-- > 0; item++);
            group = reinterpret_cast<GROUP*>(item);
        }
Fail:
        return hr;
    }

    HRESULT ReadName() throw()
    {
        const UINT32 size = this->header.namesize;
        HRESULT hr;
        FAIL_HR(hr = SafeAlloc(size, this->names));
        FAIL_HR(hr = this->StreamSeekTo(this->header.nameoff));
        FAIL_HR(this->StreamRead(this->names, size));
Fail:
        return hr;
    }

    UINT32 SizeofGroup() const throw()
    {
        return __super::SizeofGroup(this->header.groups, this->header.items);
    }

    HRESULT StreamSeekTo(UINT32 offset) const throw()
    {
        return __super::StreamSeekTo(this->stream, offset);
    }

    HRESULT StreamRead(void* data, DWORD size) const throw()
    {
        return ::StreamRead(this->stream, data, size);
    }
};

class CacheWriter : CacheBase
{
private:
    struct MemItem
    {
        UINT32 compsize;
        UINT32 orgsize;
        wchar_t* str;
        void* image;

        MemItem() throw()
            : str()
            , image()
        {
        }
        ~MemItem() throw()
        {
            SafeFree(str);
            SafeFree(image);
        }
    };

    typedef CPtrArray<MemItem, 0, TPtrArrayAutoDeleteHelper<MemItem>> ItemArray;

    struct MemGroup
    {
        ItemArray itemArray;
        wchar_t* str;

        MemGroup() throw()
            : str()
        {
        }
        ~MemGroup() throw()
        {
            SafeFree(str);
        }
    };

    typedef CPtrArray<MemGroup, 0, TPtrArrayAutoDeleteHelper<MemGroup>> GroupArray;

    GroupArray groupArray;
    UINT32 cItems;

public:
    CacheWriter() throw()
        : cItems()
    {
    }

    UINT32 GetCount() const throw()
    {
        return this->cItems;
    }

    HRESULT AddGroup(const wchar_t* name, __out UINT32* id) throw()
    {
        HRESULT hr;
        wchar_t* dup = NULL;
        MemGroup* group = NULL;
        int index;

        FAIL_HR(hr = SafeStrDupW(name, dup));
        FAIL_HR(hr = SafeNew(group));
        group->itemArray.Create();
        group->str = dup;
        this->groupArray.Create();
        index = this->groupArray.Add(group);
        FAIL_HR(hr = index != DA_ERR ? S_OK : E_OUTOFMEMORY);
        *id = static_cast<UINT32>(index);
        return hr;
Fail:
        SafeFree(dup);
        SafeDelete(group);
        return hr;
    }

    HRESULT AddImage(UINT32 group, const wchar_t* name, const void* image, UINT32 size) throw()
    {
        HRESULT hr;
        wchar_t* dup = NULL;
        MemItem* item = NULL;
        MemGroup* g = this->groupArray.GetAt(group);
        int index;
        FAIL_HR(hr = g != NULL ? S_OK : E_INVALIDARG);
        FAIL_HR(hr = SafeStrDupW(name, dup));
        FAIL_HR(hr = SafeNew(item));
        item->orgsize = size;
        item->str = dup;
        FAIL_HR(hr = Compress(image, size, &item->compsize, &item->image));
        index = g->itemArray.Add(item);
        FAIL_HR(hr = index != DA_ERR ? S_OK : E_OUTOFMEMORY);
        this->cItems++;
        return hr;
Fail:
        SafeFree(dup);
        SafeDelete(item);
        return hr;
    }

    HRESULT Write(IStream* stream) const throw()
    {
        HRESULT hr = S_OK;
        FAIL_HR(hr = this->WriteHeader(stream));
        FAIL_HR(hr = this->WriteGroup(stream));
        FAIL_HR(hr = this->WriteImage(stream));
        FAIL_HR(hr = this->WriteName(stream));
Fail:
        return hr;
    }

private:
    static HRESULT Compress(const void* src, UINT32 srcSize, __out UINT32* dstSize, __deref_out void** dstPtr) throw()
    {
        HRESULT hr;
        BYTE* dst = NULL;
        char* tmp = NULL;
        int newSize;

        FAIL_HR(hr = SafeAlloc(srcSize, tmp));
#if LZ4_COMPRESS_MODE == 1
        newSize = ::LZ4_compress_default(static_cast<const char*>(src), tmp, srcSize, srcSize);
#elif LZ4_COMPRESS_MODE == 2
        newSize = ::LZ4_compress_HC(static_cast<const char*>(src), tmp, srcSize, srcSize, LZ4HC_CLEVEL_DEFAULT);
#else
        newSize = 0;
#endif

        if (newSize > 0 && static_cast<UINT32>(newSize) < srcSize) {
            FAIL_HR(hr = SafeAlloc(newSize, dst));
            CopyMemory(dst, tmp, newSize);
            *dstSize = static_cast<UINT32>(newSize);
        } else {
            FAIL_HR(hr = SafeAlloc(srcSize, dst));
            CopyMemory(dst, src, srcSize);
            *dstSize = srcSize;
        }
        *dstPtr = dst;
        dst = NULL;
Fail:
        SafeFree(tmp);
        SafeFree(dst);
        return hr;
    }

    UINT32 SizeofGroup() const throw()
    {
        return __super::SizeofGroup(this->groupArray.GetCount(), this->cItems);
    }

    void SizeofImageAndName(UINT32* cbImage, UINT32* cbName) const throw()
    {
        const UINT32 cGroups = this->groupArray.GetCount();
        *cbImage = 0;
        *cbName = 0;
        for (UINT32 g = 0; g < cGroups; g++) {
            MemGroup* group = this->groupArray.FastGetAt(g);
            UINT32 c = group->itemArray.GetCount();
            *cbName += SizeofStringAligned(group->str);
            for (UINT32 i = 0; i < c; i++) {
                MemItem* item = group->itemArray.FastGetAt(i);
                *cbImage += AlignPtr(offsetof(IMAGE, data[item->compsize]));
                *cbName += SizeofStringAligned(item->str);
            }
        }
        *cbImage = Align16(*cbImage);
        *cbName = Align16(*cbName);
    }

    HRESULT WriteHeader(IStream* stream) const throw()
    {
        UINT32 cbImage, cbName;
        this->SizeofImageAndName(&cbImage, &cbName);
        HEADER header;
        header.signature = SIGNATURE;
        header.reserved = 0;
        header.groups = this->groupArray.GetCount();
        header.items = this->cItems;
        header.imageoff = sizeof(HEADER) + this->SizeofGroup();
        header.imagesize = cbImage;
        header.nameoff = header.imageoff + header.imagesize;
        header.namesize = cbName;
        return StreamWrite(stream, &header, sizeof(header));
    }

    HRESULT WriteGroup(IStream* stream) const throw()
    {
        HRESULT hr;
        const UINT32 size = this->SizeofGroup();
        GROUP* groups = NULL;
        FAIL_HR(hr = SafeAlloc(size, groups));
        ZeroMemory(groups, size);
        const UINT32 cGroups = this->groupArray.GetCount();
        UINT32 offImage = 0;
        UINT32 offName = 0;
        GROUP* g = groups;
        for (UINT32 i = 0; i < cGroups; i++) {
            MemGroup* group = this->groupArray.FastGetAt(i);
            UINT32 c = group->itemArray.GetCount();
            g->name = offName / 2;
            g->count = c;
            offName += SizeofStringAligned(group->str);
            g = reinterpret_cast<GROUP*>(&g->items[c]);
        }
        g = groups;
        for (UINT32 i = 0; i < cGroups; i++) {
            MemGroup* group = this->groupArray.FastGetAt(i);
            UINT32 c = group->itemArray.GetCount();
            ITEM* gi = g->items;
            for (UINT32 j = 0; j < c; j++, gi++) {
                MemItem* item = group->itemArray.FastGetAt(j);
                gi->image = offImage;
                gi->name = offName / 2;
                offImage += AlignPtr(offsetof(IMAGE, data[item->compsize]));
                offName += SizeofStringAligned(item->str);
            }
            g = reinterpret_cast<GROUP*>(gi);
        }
        FAIL_HR(hr = StreamWrite(stream, groups, size));
Fail:
        SafeFree(groups);
        return hr;
    }

    HRESULT WriteImage(IStream* stream) const throw()
    {
        HRESULT hr = S_OK;
        const UINT32 cGroups = this->groupArray.GetCount();
        UINT32 size = 0;
        for (UINT32 i = 0; i < cGroups; i++) {
            MemGroup* group = this->groupArray.FastGetAt(i);
            UINT32 c = group->itemArray.GetCount();
            for (UINT32 j = 0; j < c; j++) {
                MemItem* item = group->itemArray.FastGetAt(j);
                IMAGE image;
                image.compsize = item->compsize;
                image.orgsize = item->orgsize;
                FAIL_HR(hr = StreamWrite(stream, &image, offsetof(IMAGE, data)));
                FAIL_HR(hr = StreamWrite(stream, item->image, item->compsize));
                FAIL_HR(hr = StreamPadWrite(stream, item->compsize));
                size += AlignPtr(offsetof(IMAGE, data[item->compsize]));
            }
        }
        hr = StreamPadWrite16(stream, size);
Fail:
        return hr;
    }

    HRESULT WriteName(IStream* stream) const throw()
    {
        HRESULT hr = S_OK;
        const UINT32 cGroups = this->groupArray.GetCount();
        UINT32 size = 0;
        for (UINT32 i = 0; i < cGroups; i++) {
            MemGroup* group = this->groupArray.FastGetAt(i);
            FAIL_HR(hr = StreamWriteStr(stream, group->str));
        }
        for (UINT32 i = 0; i < cGroups; i++) {
            MemGroup* group = this->groupArray.FastGetAt(i);
            UINT32 c = group->itemArray.GetCount();
            for (UINT32 j = 0; j < c; j++) {
                MemItem* item = group->itemArray.FastGetAt(j);
                FAIL_HR(hr = StreamWriteStr(stream, item->str));
                size += SizeofStringAligned(item->str);
            }
        }
        hr = StreamPadWrite16(stream, size);
Fail:
        return hr;
    }
};


class DibImage
{
private:
    UINT32* data;
    UINT width;
    UINT height;

    static HRESULT WICLoadImage(
        __in IWICImagingFactory* factory,
        __in const wchar_t* path,
        __deref_out IWICBitmapFrameDecode** outPtr) throw()
    {
        *outPtr = NULL;
        HRESULT hr;
        ComPtr<IWICBitmapDecoder> decoder;

        FAIL_HR(hr = factory->CreateDecoderFromFilename(path, NULL, GENERIC_READ, WICDecodeMetadataCacheOnLoad, &decoder));
        FAIL_HR(hr = decoder->GetFrame(0, outPtr));

Fail:
        return hr;
    }

    static HRESULT WICConvertBitmapSource(
        __in IWICImagingFactory* factory,
        __in REFWICPixelFormatGUID format,
        __in IWICBitmapSource* bmp,
        __deref_out IWICBitmapSource** outPtr) throw()
    {
        *outPtr = NULL;
        HRESULT hr;
        ComPtr<IWICFormatConverter> converter;
        FAIL_HR(hr = factory->CreateFormatConverter(&converter));
        FAIL_HR(hr = converter->Initialize(
                        bmp, format, WICBitmapDitherTypeNone, NULL, 0,
                        WICBitmapPaletteTypeMedianCut));
        *outPtr = converter.Detach();

Fail:
        return hr;
    }

    static HRESULT WICLoadImage32(
        __in IWICImagingFactory* factory,
        __in const wchar_t* path,
        __deref_out IWICBitmapSource** outPtr) throw()
    {
        *outPtr = NULL;
        HRESULT hr;
        ComPtr<IWICBitmapFrameDecode> bmp;
        FAIL_HR(hr = WICLoadImage(factory, path, &bmp));
        FAIL_HR(hr = WICConvertBitmapSource(factory, GUID_WICPixelFormat32bppBGRA, bmp, outPtr));

Fail:
        return hr;
    }

public:
    DibImage() throw()
        : data()
        , width()
        , height()
    {
    }

    ~DibImage() throw()
    {
        SafeFreeAligned(this->data);
    }

    HRESULT Init(IWICImagingFactory* factory, __in const wchar_t* path) throw()
    {
        HRESULT hr;
        ComPtr<IWICBitmapSource> bmp;
        UINT cx = 0, cy = 0, stride = 0, size = 0;
        BYTE* bits = NULL;

        FAIL_HR(hr = WICLoadImage32(factory, path, &bmp));
        FAIL_HR(hr = bmp->GetSize(&cx, &cy));

        FAIL_HR(hr = cx < 65536 && cy < 65536 ? S_OK : E_FAIL);
        // the width of an image must be the multiple of 8
        FAIL_HR(hr = (cx % 8) == 0 ? S_OK : E_NOTIMPL);
        stride = cx * 4;
        size = stride * cy;

        FAIL_HR(hr = SafeAllocAligned(size, bits));
        ZeroMemory(bits, size);

        FAIL_HR(hr = bmp->CopyPixels(NULL, stride, size, bits));

        this->data = reinterpret_cast<UINT32*>(bits);
        bits = NULL;

Fail:
        SafeFreeAligned(bits);
        this->width = cx;
        this->height = cy;
        return hr;
    }

    UINT GetWidth() const throw()
    {
        return this->width;
    }

    UINT GetHeight() const throw()
    {
        return this->height;
    }

    UINT32* GetBits() const throw()
    {
        return this->data;
    }

#define INIT \
    __m128i zero = _mm_setzero_si128(), \
            ones = _mm_cmpeq_epi32(zero, zero), \
            alphas = _mm_slli_epi32(ones, 24)

#define CROUCH(n)   __m128i x##n = _mm_load_si128(reinterpret_cast<const __m128i*>(src + (n * 4)))
#define BIND(n)     __m128i y##n = _mm_andnot_si128(_mm_cmpeq_epi32(_mm_and_si128(x##n, alphas), zero), x##n)
#define SET(n)      _mm_stream_si128(reinterpret_cast<__m128i*>(dst + (n * 4)), y##n)

    void CopyImage(UINT32* dst, UINT dstX, UINT dstY, UINT dstWidth) const throw()
    {
        ASSERT((dstX % 4) == 0);
        const UINT32* src = this->data;
        const UINT cx = this->width;
        const UINT cy = this->height;
        dst += dstWidth * dstY + dstX;
        INIT;
        for (UINT y = 0; y < cy; y++) {
            for (UINT x = 0; x < cx; x += 8) {
                CROUCH(0); CROUCH(1);
                BIND(0); BIND(1);
                SET(0); SET(1);
                src += 8;
                dst += 8;
            }
            dst += dstWidth - cx;
        }
    }

#undef SET
#undef BIND
#undef CROUCH
#undef INIT

#define INIT        __m128i zero = _mm_setzero_si128()

#define CROUCH(n)   __m128i x##n = _mm_load_si128(reinterpret_cast<const __m128i*>(src + (n * 4)))
#define BIND(n)     __m128i y##n = _mm_shuffle_epi32(_mm_srli_epi32(x##n, 24), _MM_SHUFFLE(0, 1, 2, 3))
#define SET(m, n)   *dst = static_cast<BYTE>(_mm_movemask_epi8(_mm_cmpeq_epi8(_mm_packs_epi16(_mm_packs_epi32(y##n, y##m), zero), zero)))

    void CopyMask(BYTE* dst) const throw()
    {
        const UINT32* src = this->data;
        const UINT cx = this->width;
        const UINT cy = this->height;
        const UINT padUp = -static_cast<int>((cx + 7) / 8) & 3;
        INIT;
        for (UINT y = 0; y < cy; y++) {
            for (UINT x = 0; x < cx; x += 8) {
                CROUCH(0); CROUCH(1);
                BIND(0); BIND(1);
                SET(0, 1);
                src += 8;
                dst++;
            }
            dst += padUp;
        }
    }

#undef SET
#undef BIND
#undef CROUCH
#undef INIT

#define TICK(p, n)      __m128i p##n = _mm_load_si128(reinterpret_cast<const __m128i*>(p + (n * 4)))
#define TOCK(s, d, n)   _mm_stream_si128(reinterpret_cast<__m128i*>(d + (n * 4)), s##n)

    void FlipVertical() throw()
    {
        const UINT cx = this->width;
        const UINT cy = this->height;
        UINT32* f = this->data;
        UINT32* l = this->data + cx * (cy - 1);
        while (f < l) {
            for (UINT x = 0; x < cx; x += 8) {
                TICK(f, 0); TICK(f, 1);
                TICK(l, 0); TICK(l, 1);
                TOCK(f, l, 0); TOCK(f, l, 1);
                TOCK(l, f, 0); TOCK(l, f, 1);
                f += 8;
                l += 8;
            }
            l -= cx * 2;
        }
    }

#undef TOCK
#undef TICK
};


class WicLoader
{
private:
    ComPtr<IWICImagingFactory> wicFactory;

protected:
    WicLoader() throw()
    {
    }

    ~WicLoader() throw()
    {
    }

    HRESULT GetWICFactory(__deref_out IWICImagingFactory** factoryPtr) throw()
    {
        HRESULT hr = S_OK;
        if (this->wicFactory == NULL) {
            hr = this->wicFactory.CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER);
        }
        this->wicFactory.CopyTo(factoryPtr);
        return hr;
    }

    HRESULT WICLoadImage(__in const wchar_t* path, __deref_out IWICBitmapFrameDecode** outPtr) throw()
    {
        *outPtr = NULL;
        HRESULT hr;
        ComPtr<IWICImagingFactory> factory;
        ComPtr<IWICBitmapDecoder> decoder;

        FAIL_HR(hr = this->GetWICFactory(&factory));
        FAIL_HR(hr = factory->CreateDecoderFromFilename(path, NULL, GENERIC_READ, WICDecodeMetadataCacheOnLoad, &decoder));
        FAIL_HR(hr = decoder->GetFrame(0, outPtr));

Fail:
        return hr;
    }
};


class Scanner : public WicLoader
{
private:
    CacheWriter cache;
    UINT32* image;
    wchar_t root32[MAX_PATH];
    wchar_t root16[MAX_PATH];

public:
    Scanner() throw()
        : image()
    {
    }

    ~Scanner() throw()
    {
        SafeFreeAligned(this->image);
    }

    void Scan(const wchar_t* root) throw()
    {
        ::GetFullPathNameW(root, _countof(this->root32), this->root32, NULL);
        ::PathCombineW(this->root16, this->root32, L"FatCow_Icons16x16_Color");
        ::PathAppendW(this->root32, L"FatCow_Icons32x32_Color");
        this->ScanDirs();
    }

    HRESULT SaveCache(const wchar_t* file) throw()
    {
        ComPtr<IStream> stream;
        HRESULT hr = ::SHCreateStreamOnFileEx(file, STGM_CREATE | STGM_WRITE | STGM_SHARE_EXCLUSIVE, 0, TRUE, NULL, &stream);
        if (SUCCEEDED(hr)) {
            hr = this->cache.Write(stream);
        }
        return hr;
    }

private:
    HRESULT ScanFiles(const wchar_t* dir) throw()
    {
        UINT32 group;
        HRESULT hr = this->cache.AddGroup(dir, &group);
        if (FAILED(hr)) {
            return hr;
        }

        WCHAR parent32[MAX_PATH], parent16[MAX_PATH], path32[MAX_PATH], path16[MAX_PATH];
        ::PathCombineW(parent32, this->root32, dir);
        ::PathCombineW(parent16, this->root16, dir);
        ::PathCombineW(path32, parent32, L"*.png");

        WIN32_FIND_DATAW wfd;
        HANDLE hfindChild = ::FindFirstFileW(path32, &wfd);
        if (hfindChild != INVALID_HANDLE_VALUE) {
            do {
                if (PathIsDotOrDotDot(wfd.cFileName) || (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
                    continue;
                }
                //if(this->cache.GetCount()>100)break;
                if ((this->cache.GetCount() % 100) == 0) {
                    TRACE("%d ", this->cache.GetCount());
                }
                wchar_t* name = wfd.cFileName;
                ::PathCombineW(path32, parent32, name);
                ::PathCombineW(path16, parent16, name);
                ::PathRemoveExtensionW(name);

                hr = this->AddImage(group, name, path32, path16);
                if (FAILED(hr)) {
                    TRACE("\nFAILED %ls/%ls\n", dir, name);
                }
            } while (::FindNextFileW(hfindChild, &wfd));
            ::FindClose(hfindChild);
        }
        return S_OK;
    }

    HRESULT AddImage(UINT32 group, const wchar_t* name, const wchar_t* path32, const wchar_t* path16) throw()
    {
        HRESULT hr;
        DibImage image32, image16;
        ComPtr<IWICImagingFactory> factory;
        const UINT32 size = IMGWIDTH * IMGHEIGHT * 4;

        if (this->image == NULL) {
            FAIL_HR(hr = SafeAllocAligned(size, this->image));
            ZeroMemory(this->image, size);
        }

        FAIL_HR(hr = this->GetWICFactory(&factory));
        FAIL_HR(hr = image32.Init(factory, path32));
        FAIL_HR(hr = image16.Init(factory, path16));

        FAIL_HR(hr = image32.GetWidth() == 32 && image32.GetHeight() == 32 ? S_OK : E_FAIL);
        FAIL_HR(hr = image16.GetWidth() == 16 && image16.GetHeight() == 16 ? S_OK : E_FAIL);

        image32.CopyImage(this->image, PADDING, PADDING, IMGWIDTH);
        image16.CopyImage(this->image, PADDING * 2 + SIZE32, PADDING + (SIZE32 - SIZE16), IMGWIDTH);

        FAIL_HR(hr = this->cache.AddImage(group, name, this->image, size));

Fail:
        return hr;
    }

    void ScanDirs() throw()
    {
        WCHAR path[MAX_PATH];
        ::PathCombineW(path, this->root32, L"*");

        WIN32_FIND_DATAW wfd;
        HANDLE hfindParent = ::FindFirstFileW(path, &wfd);
        if (hfindParent != INVALID_HANDLE_VALUE) {
            do {
                if (PathIsDotOrDotDot(wfd.cFileName) || (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
                    continue;
                }
                this->ScanFiles(wfd.cFileName);
            } while (::FindNextFileW(hfindParent, &wfd));
            ::FindClose(hfindParent);
        }
        TRACE(L"%d\n", this->cache.GetCount());
    }

    static bool PathIsDotOrDotDot(__in const wchar_t* path) throw()
    {
        return path[0] == L'.' && (path[1] == L'\0' || (path[1] == L'.' && path[2] == L'\0'));
    }
};

class FolderList
{
public:
    static const int MAX_ITEMS = 10;

private:
    typedef CPtrArray<wchar_t, 0, TPtrArrayAutoFreeHelper<wchar_t>> ItemArray;
    ItemArray items;

    static bool IsSpaceChar(__in BYTE c) throw()
    {
        return (c == ' ') || (c == '\t');
    }

    static BYTE* SkipWhitespaces(__in BYTE* s) throw()
    {
        for (; IsSpaceChar(*s); s++);
        return s;
    }

    static BYTE* ReadData(__in LPCWSTR path, __out DWORD* fileSize) throw()
    {
        HRESULT hr;
        ComPtr<IStream> stream;
        BYTE* data = NULL;
        ULARGE_INTEGER size;
        ULONG read;

        FAIL_HR(hr = ::SHCreateStreamOnFileEx(path, STGM_READ | STGM_SHARE_DENY_WRITE, 0, FALSE, NULL, &stream));

        FAIL_HR(hr = ::IStream_Size(stream, &size));
        FAIL_HR(hr = size.HighPart == 0 && size.LowPart <= 0xFFFFFFF0U ? S_OK : E_OUTOFMEMORY);
        FAIL_HR(hr = SafeAlloc(size.LowPart + 1, data));
        FAIL_HR(hr = stream->Read(data, size.LowPart + 1, &read));
        FAIL_HR(hr = read > 0 && read <= size.LowPart ? S_OK : E_FAIL);
        data[read] = 0;
        *fileSize = read;
        return data;

Fail:
        SafeFree(data);
        *fileSize = 0;
        return NULL;
    }

    static char* StrToUtf8(const wchar_t* src) throw()
    {
        int cch = ::WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, src, -1, NULL, 0, NULL, NULL);
        if (cch <= 0) {
            return NULL;
        }
        char* dst;
        SafeAlloc(cch, dst);
        if (dst == NULL) {
            return NULL;
        }
        int copied = ::WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, src, -1, dst, cch, NULL, NULL);
        if (copied > 0 && copied <= cch) {
            return dst;
        }
        SafeFree(dst);
        return NULL;
    }

    static wchar_t* Utf8ToStr(const char* src) throw()
    {
        int cch = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, src, -1, NULL, 0);
        if (cch <= 0) {
            return NULL;
        }
        wchar_t* dst;
        SafeAlloc(cch * 2, dst);
        if (dst == NULL) {
            return NULL;
        }
        int copied = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, src, -1, dst, cch);
        if (copied > 0 && copied <= cch) {
            return dst;
        }
        SafeFree(dst);
        return NULL;
    }

    static int CALLBACK sCompare(void* p1, void* p2, LPARAM) throw()
    {
        return ::CompareStringOrdinal(static_cast<LPWSTR>(p1), -1, static_cast<LPWSTR>(p2), -1, TRUE) - CSTR_EQUAL;
    }

    HRESULT AddWorker(const wchar_t* path, int at, int* i) throw()
    {
        *i = -1;
        if (path == NULL || *path == L'\0') {
            return S_FALSE;
        }
        if (this->items.GetCount() >= MAX_ITEMS) {
            this->items.RemoveAt(MAX_ITEMS);
        }
        if (this->items.Search(const_cast<LPWSTR>(path), 0, sCompare, 0, DPAS_SORTED) != -1) {
            return S_FALSE;
        }
        LPWSTR dup;
        HRESULT hr = SafeStrDupW(path, dup);
        if (SUCCEEDED(hr)) {
            *i = this->items.InsertAt(at, dup);
            hr = *i != DA_ERR ? S_OK : E_OUTOFMEMORY;
        }
        if (SUCCEEDED(hr)) {
            dup = NULL;
        }
        SafeFree(dup);
        return hr;
    }

    void ReadList(BYTE* data) throw()
    {
        if (data[0] == 0xef && data[1] == 0xbb && data[2] == 0xbf) {
            data += 3;
        }
        for (BYTE* cur = data; *cur != 0; ) {
            cur = SkipWhitespaces(cur);
            if (*cur == 0) {
                break;
            }
            BYTE* eol;
            BYTE* next;
            for (eol = cur + 1; (*eol != 0) && ((*eol != '\r') && (*eol != '\n')); eol++);
            for (next = eol + 1; (*next == '\r') || (*next == '\n'); next++);

            if (*cur != ';') {
                for (eol--; (eol > cur) && IsSpaceChar(*eol); eol--);
                *++eol = 0;
                wchar_t* str = Utf8ToStr(reinterpret_cast<const char*>(cur));
                int i;
                this->AddWorker(str, DA_LAST, &i);
                SafeFree(str);
            }
            cur = next;
        }
    }

    HRESULT SaveList(IStream* stream) throw()
    {
        HRESULT hr;
        ULARGE_INTEGER zero = { 0, 0 };
        const BYTE bom[] = { 0xef, 0xbb, 0xbf };
        const BYTE crlf[] = { '\r', '\n' };
        char* str = NULL;
        FAIL_HR(hr = stream->SetSize(zero));
        FAIL_HR(hr = StreamWrite(stream, bom, 3));
        const int cItems = this->GetCount();
        for (int i = 0; i < cItems; i++) {
            str = StrToUtf8(this->GetAt(i));
            if (str != NULL) {
                FAIL_HR(hr = StreamWrite(stream, str, strlen(str)));
                FAIL_HR(hr = StreamWrite(stream, crlf, 2));
            }
            SafeFree(str);
        }
Fail:
        SafeFree(str);
        return hr;
    }

public:
    FolderList() throw()
    {
    }

    ~FolderList() throw()
    {
    }

    HRESULT Load(__in LPCWSTR fileName) throw()
    {
        HRESULT hr;
        FAIL_HR(hr = this->items.Create() ? S_OK : E_OUTOFMEMORY);
        DWORD cb;
        BYTE* ptr = ReadData(fileName, &cb);
        if (ptr != NULL) {
            this->ReadList(ptr);
        }
        SafeFree(ptr);
        return S_OK;
Fail:
        return hr;
    }

    void Save(__in LPCWSTR fileName) throw()
    {
        if (fileName[0] == L'\0') {
            return;
        }
        ComPtr<IStream> stream;
        HRESULT hr = ::SHCreateStreamOnFileEx(fileName, STGM_CREATE | STGM_WRITE | STGM_SHARE_EXCLUSIVE, 0, TRUE, NULL, &stream);
        if (SUCCEEDED(hr)) {
            hr = this->SaveList(stream);
        }
    }

    HRESULT Add(LPCWSTR path, int* i) throw()
    {
        return this->AddWorker(path, 0, i);
    }

    int GetCount() const throw()
    {
        return this->items.GetCount();
    }

    LPCWSTR GetAt(__in int i) const throw()
    {
        return this->items.GetAt(i);
    }
};

enum ControlID
{
    IDC_LISTVIEW = 1,
    IDC_STATUSBAR,
    IDC_EDIT,
    IDC_LAST,
};

HBITMAP Create32bppTopDownDIB(int cx, int cy, __deref_out_opt void** bits) throw()
{
    BITMAPINFOHEADER bmi = {
        sizeof(BITMAPINFOHEADER),
        cx, -cy, 1, 32
    };
    return ::CreateDIBSection(NULL, reinterpret_cast<BITMAPINFO*>(&bmi), DIB_RGB_COLORS, bits, NULL, 0);
}

BOOL GetWorkAreaRect(HWND hwnd, __out RECT* prc) throw()
{
    HMONITOR hMonitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    if (hMonitor != NULL) {
        MONITORINFO mi = { sizeof(MONITORINFO) };
        BOOL bRet = GetMonitorInfo(hMonitor, &mi);
        if (bRet) {
            *prc = mi.rcWork;
            return bRet;
        }
    }
    return FALSE;
}

void SetWindowPosition(HWND hwnd, int x, int y, int cx, int cy) throw()
{
    RECT rcWork;
    if (GetWorkAreaRect(hwnd, &rcWork)) {
        int dx = 0, dy = 0;
        if (x < rcWork.left) {
            dx = rcWork.left - x;
        }
        if (y < rcWork.top) {
            dy = rcWork.top - y;
        }
        if (x + cx + dx > rcWork.right) {
            dx = rcWork.right - x - cx;
            if (x + dx < rcWork.left) {
                dx += rcWork.left - x;
            }
        }
        if (y + cy + dy > rcWork.bottom) {
            dy = rcWork.bottom - y - cy;
            if (y + dx < rcWork.top) {
                dy += rcWork.top - y;
            }
        }
        ::MoveWindow(hwnd, x + dx, y + dy, cx, cy, TRUE);
    }
}

void MyAdjustWindowRect(HWND hwnd, __inout LPRECT lprc) throw()
{
    ::AdjustWindowRect(lprc, WS_OVERLAPPEDWINDOW, TRUE);
    RECT rect = *lprc;
    rect.bottom = 0x7fff;
    FORWARD_WM_NCCALCSIZE(hwnd, FALSE, &rect, ::SendMessage);
    lprc->bottom += rect.top;
}

HMENU GetSubMenuById(HMENU hm, UINT id)
{
    MENUITEMINFO mi = { sizeof(mi), MIIM_SUBMENU };
    if (::GetMenuItemInfo(hm, id, false, &mi) == FALSE) {
        return NULL;
    }
    return mi.hSubMenu;
}

bool SetSubMenuById(HMENU hm, UINT id, HMENU hsub)
{
    MENUITEMINFO mi = { sizeof(mi), MIIM_SUBMENU };
    mi.hSubMenu = hsub;
    return ::SetMenuItemInfo(hm, id, false, &mi) != FALSE;
}

HRESULT BrowseForFolder(__in_opt HWND hwndOwner, __deref_out LPWSTR* path) throw()
{
    *path = NULL;
    HRESULT hr;
    ComPtr<IFileOpenDialog> dlg;
    ComPtr<IShellItem> si;
    DWORD options;

    FAIL_HR(hr = dlg.CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER));
    FAIL_HR(hr = dlg->GetOptions(&options));
    FAIL_HR(hr = dlg->SetOptions(options | FOS_FORCEFILESYSTEM | FOS_PICKFOLDERS));
    FAIL_HR(hr = dlg->Show(hwndOwner));
    FAIL_HR(hr = dlg->GetResult(&si));

    FAIL_HR(hr = si->GetDisplayName(SIGDN_FILESYSPATH, path));

Fail:
    return hr;
}


template <class T>
class BaseWindow
{
private:
    static LRESULT CALLBACK sWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) throw()
    {
        LPCREATESTRUCT lpcs = reinterpret_cast<LPCREATESTRUCT>(lParam);
        T* self;
        if (message == WM_NCCREATE && lpcs != NULL) {
            self = static_cast<T*>(lpcs->lpCreateParams);
            ::SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        } else {
            self = reinterpret_cast<T*>(::GetWindowLongPtr(hwnd, GWLP_USERDATA));
        }
        if (self != NULL) {
            return self->WindowProc(hwnd, message, wParam, lParam);
        } else {
            return ::DefWindowProc(hwnd, message, wParam, lParam);
        }
    }

protected:
    static ATOM Register(LPCTSTR className, LPCTSTR menu = NULL) throw()
    {
        WNDCLASS wc = {
            0, sWindowProc, 0, 0,
            HINST_THISCOMPONENT,
            NULL, ::LoadCursor(NULL, IDC_ARROW),
            0, menu,
            className };
        return ::RegisterClass(&wc);
    }

    HWND Create(ATOM atom, LPCTSTR name) throw()
    {
        return ::CreateWindow(MAKEINTATOM(atom), name,
            WS_VISIBLE | WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT, CW_USEDEFAULT,
            CW_USEDEFAULT, CW_USEDEFAULT,
            NULL, NULL, HINST_THISCOMPONENT, this);
    }
};

class SearchEdit
{
private:
    HFONT           font;
    bool            inSize;

    static LRESULT CALLBACK sWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR ref) throw()
    {
        SearchEdit* self = reinterpret_cast<SearchEdit*>(ref);
        return self->WindowProc(hwnd, message, wParam, lParam);
    }

    LRESULT WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) throw()
    {
        switch (message) {
        case WM_DESTROY:
            ::RemoveWindowSubclass(hwnd, sWindowProc, 0);
            break;

        case WM_SIZE:
            this->OnSize(hwnd);
            break;

        case WM_CHAR:
            if (this->OnChar(hwnd, wParam)) {
                return 0;
            }
            break;
        }
        return ::DefSubclassProc(hwnd, message, wParam, lParam);
    }

    void OnSize(HWND hwnd) throw()
    {
        if (this->inSize) {
            return;
        }
        int height = GetFontHeight(hwnd);
        RECT rect = { 0, 0, 1, height + 2 };
        ::AdjustWindowRectEx(&rect, GetWindowStyle(hwnd), false, GetWindowExStyle(hwnd));
        int y = rect.bottom - rect.top;
        ::GetClientRect(::GetParent(hwnd), &rect);
        this->inSize = true;
        ::MoveWindow(hwnd, 0, 0, rect.right, y, TRUE);
        this->inSize = false;
    }

    bool OnChar(HWND hwnd, WPARAM wParam) throw()
    {
        if (wParam == VK_ESCAPE) {
            Edit_SetText(hwnd, NULL);
            return true;
        }
        if (wParam == VK_RETURN) {
            return true;
        }
        return false;
    }

    static int GetFontHeight(HWND hwnd) throw()
    {
        TEXTMETRIC tm;
        tm.tmHeight = 0;
        HDC hdc = ::CreateCompatibleDC(NULL);
        if (hdc != NULL) {
            HFONT fontOld = SelectFont(hdc, GetWindowFont(hwnd));
            ::GetTextMetrics(hdc, &tm);
            SelectFont(hdc, fontOld);
            ::DeleteDC(hdc);
        }
        return tm.tmHeight > 0 ? tm.tmHeight : 16;
    }

public:
    SearchEdit() throw()
        : font()
        , inSize()
    {
    }

    ~SearchEdit() throw()
    {
        if (this->font != NULL) {
            DeleteFont(this->font);
        }
    }

    HWND Create(HWND hwndParent, int id) throw()
    {
        HWND hwnd = ::CreateWindowEx(WS_EX_CLIENTEDGE, WC_EDIT, NULL,
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
            | ES_AUTOHSCROLL,
            0, 0, 0, 0,
            hwndParent,
            reinterpret_cast<HMENU>(id),
            HINST_THISCOMPONENT, NULL);
        if (hwnd != NULL) {
            NONCLIENTMETRICS ncm = { sizeof(ncm) };
            ::SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
            this->font = ::CreateFontIndirect(&ncm.lfStatusFont);
            SetWindowFont(hwnd, this->font, FALSE);
            ::SetWindowSubclass(hwnd, sWindowProc, 0, reinterpret_cast<DWORD_PTR>(this));
            Edit_SetCueBannerTextFocused(hwnd, L"Search", true);
        }
        return hwnd;
    }
};

class StatusBar
{
public:
    enum Parts
    {
        PART_STAT,
        PART_NAME,
        PART_GROUP,
        PART_SEL,
        PART_ALL,
        PART_LAST,
    };

private:
    static LRESULT CALLBACK sWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) throw()
    {
        LRESULT lr = ::DefSubclassProc(hwnd, message, wParam, lParam);
        switch (message) {
        case WM_DESTROY:
            ::RemoveWindowSubclass(hwnd, sWindowProc, 0);
            break;

        case WM_SIZE:
            OnSize(hwnd);
            break;
        }
        return lr;
    }

    static void OnSize(HWND hwnd) throw()
    {
        TEXTMETRIC tm;
        tm.tmAveCharWidth = 0;
        HDC hdc = ::GetDC(hwnd);
        if (hdc != NULL) {
            ::GetTextMetrics(hdc, &tm);
            ::ReleaseDC(hwnd, hdc);
        }
        RECT rect;
        ::GetClientRect(hwnd, &rect);
        const int cx = tm.tmAveCharWidth;
        int x = rect.right;
        int widths[PART_LAST];
        widths[PART_ALL] = x;
        x -= cx * 20;
        widths[PART_SEL] = x;
        x -= cx * 20;
        widths[PART_GROUP] = x;
        x -= cx * 24;
        widths[PART_NAME] = x;
        x -= cx * 40;
        widths[PART_STAT] = x;
        ::SendMessage(hwnd, SB_SETPARTS, PART_LAST, reinterpret_cast<LPARAM>(widths));
    }

public:
    static HWND Create(HWND hwndParent, int id) throw()
    {
        HWND hwnd = ::CreateStatusWindow(WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | SBARS_SIZEGRIP, NULL, hwndParent, id);
        if (hwnd != NULL) {
            ::SetWindowSubclass(hwnd, sWindowProc, 0, 0);
        }
        return hwnd;
    }

    static bool SetText(HWND hwnd, Parts part, LPCWSTR text) throw()
    {
        return ::SendMessage(hwnd, SB_SETTEXTW, part, reinterpret_cast<LPARAM>(text)) != FALSE;
    }
};

class FolderMenu
{
private:
    HMENU           popup;
    UINT            firstId;

public:
    FolderMenu() throw()
        : popup()
        , firstId()
    {
    }

    ~FolderMenu() throw()
    {
    }

    HMENU GetMenu() const throw()
    {
        return this->popup;
    }

    void Init(HMENU hm, UINT menuId, UINT itemId) throw()
    {
        this->popup = GetSubMenuById(hm, menuId);
        this->firstId = itemId;
        TRACE("%p, %u\n", this->popup, this->firstId);
    }

    void HandleInitMenuPopup(HMENU hm, const FolderList& list) throw()
    {
        if (this->popup == hm) {
            while (::GetMenuItemCount(hm) > 2) {
                ::DeleteMenu(hm, 0, MF_BYPOSITION);
            }
            MENUITEMINFOW mi = { sizeof(mi), MIIM_TYPE | MIIM_ID, MFT_STRING };
            const int cItems = list.GetCount();
            for (int i = 0; i < cItems; i++) {
                mi.wID = this->firstId + i + 1;
                mi.dwTypeData = const_cast<LPWSTR>(list.GetAt(i));
                ::InsertMenuItemW(hm, i, true, &mi);
            }
        }
    }
};

#include "res.rc"

class MainWindow : public BaseWindow<MainWindow>, public WicLoader
{
protected:
    friend BaseWindow<MainWindow>;

#include <pshpack1.h>

    struct LazyIconEntry
    {
        UINT8 width;
        UINT8 height;
        UINT8 zero1;
        UINT8 zero2;
        UINT16 one;
        UINT16 thirtytwo;
        UINT32 size;
        UINT32 offset;
    };

    struct LazyIconDir
    {
        UINT16 zero;
        UINT16 one;
        UINT16 two;
        LazyIconEntry entries[2];
    };

#include <poppack.h>

    struct MemItem
    {
        int image;
        UINT32 group;
        UINT32 item;
    };

    typedef CPtrArray<MemItem, 256, TPtrArrayDestroyHelper<MemItem>> ItemArray;
    typedef HRESULT (MainWindow::*COPYCOMMAND)(HWND, LPCWSTR, LPCWSTR, LPCWSTR, LPCWSTR);

    Cache           cache;
    SearchEdit      search;
    ItemArray       filter;
    FolderList      folders;
    FolderMenu      pngMenu;
    FolderMenu      bmpMenu;
    FolderMenu      icoMenu;

    HWND            edit;
    HWND            listView;
    HWND            status;
    HMENU           popup;
    HBITMAP         bitmap;
    void*           bits;
    ComPtr<IImageList2> imageList;
    bool            shown;
    MemItem*        items;
    UINT            cItems;
    INT             ctlArray[2 * IDC_LAST];
    WCHAR           appPath[MAX_PATH];
    WCHAR           cachePath[MAX_PATH];
    WCHAR           listPath[MAX_PATH];

    HRESULT LoadIcons(HWND hwnd) throw()
    {
        HRESULT hr;

        ::GetModuleFileNameW(NULL, this->appPath, _countof(this->appPath));
        ::PathRemoveFileSpecW(this->appPath);
        ::PathCombineW(this->cachePath, this->appPath, L"cache.db");
        ::PathCombineW(this->listPath, this->appPath, L"folders.lst");

        if (::PathFileExistsW(this->cachePath) == FALSE) {
            StatusBar::SetText(this->status, StatusBar::PART_STAT, L"Building cache...");
            Scanner scanner;
            scanner.Scan(this->appPath);
            scanner.SaveCache(this->cachePath);
        }

        StatusBar::SetText(this->status, StatusBar::PART_STAT, L"Loading icons...");

        hr = this->LoadCache();
        FAIL_HR(hr);
        FAIL_HR(hr = this->LoadList());

        FAIL_HR(hr = ::ImageList_CoCreateInstance(CLSID_ImageList, NULL, IID_PPV_ARGS(&this->imageList)));
        FAIL_HR(hr = this->imageList->Initialize(IMGWIDTH, IMGHEIGHT, ILC_COLOR32 | ILC_MASK, 0, 0));

        ListView_SetImageList(this->listView, IImageListToHIMAGELIST(this->imageList), LVSIL_NORMAL);

        this->PopulateItems(hwnd);
Fail:

        StatusBar::SetText(this->status, StatusBar::PART_STAT, NULL);
        return hr;
    }

    HRESULT LoadCache() throw()
    {
        HRESULT hr;
        ComPtr<IStream> stream;
        FAIL_HR(hr = ::SHCreateStreamOnFileEx(this->cachePath, STGM_READ | STGM_SHARE_DENY_WRITE, 0, FALSE, NULL, &stream));

        FAIL_HR(hr = this->cache.Init(stream));
        FAIL_HR(hr = this->LoadCacheWorker());

Fail:
        return hr;
    }

    HRESULT LoadCacheWorker() throw()
    {
        const UINT count = this->cache.GetAllItemCount();
        this->items = new MemItem[count];
        if (this->items == NULL) {
            return E_OUTOFMEMORY;
        }

        if (this->filter.Create() == false) {
            return E_OUTOFMEMORY;
        }

        UINT index = 0;
        const UINT32 cGroups = this->cache.GetGroupCount();
        for (UINT32 i = 0; i < cGroups; i++) {
            const UINT32 c = this->cache.GetItemCount(i);
            for (UINT32 j = 0; j < c; j++) {
                if (index >= count) {
                    this->cItems = index;
                    return S_FALSE;
                }
                MemItem& item = this->items[index++];
                item.image = -1;
                item.group = i;
                item.item = j;
            }
        }
        this->cItems = index;
        return S_OK;
    }

    HRESULT LoadList() throw()
    {
        return this->folders.Load(listPath);
    }

    void PopulateItems(HWND hwnd) throw()
    {
        wchar_t text[256];
        Edit_GetText(this->edit, text, _countof(text));
        wchar_t* query = text[0] != L'\0' ? text : NULL;

        this->filter.RemoveAll();

        const bool regex = query != NULL && ::wcspbrk(query, L"^$.?*+") != NULL;
        for (UINT32 i = 0, count = this->cItems; i < count; i++) {
            MemItem* item = &this->items[i];
            if (query != NULL) {
                const wchar_t* name = this->cache.GetItemName(item->group, item->item);
                const wchar_t* matched = regex ? ::match(name, query) : ::StrStrIW(name, query);
                if (matched == NULL) {
                    continue;
                }
            }
            this->filter.Add(item);
        }

        const int listCount = this->filter.GetCount();
        SetWindowRedraw(this->listView, false);
        ListView_SetItemCount(this->listView, listCount);
        ListView_SetItemState(this->listView, 0, LVIS_FOCUSED, LVIS_FOCUSED);
        ListView_EnsureVisible(this->listView, 0, false);
        SetWindowRedraw(this->listView, true);

        ::StringCchPrintfW(text, _countof(text), L"%d icons", listCount);
        StatusBar::SetText(this->status, StatusBar::PART_ALL, text);
        this->UpdateStatus(hwnd);
    }

    void UpdateStatus(HWND hwnd) throw()
    {
        UINT count = ListView_GetSelectedCount(this->listView);
        wchar_t text[64];
        if (count > 0) {
            ::StringCchPrintfW(text, _countof(text), L"%d selected", count);
        } else {
            *text = L'\0';
        }
        StatusBar::SetText(this->status, StatusBar::PART_NAME, NULL);
        StatusBar::SetText(this->status, StatusBar::PART_SEL, text);

        int index = ListView_GetNextItem(this->listView, -1, LVNI_SELECTED | LVNI_FOCUSED);
        if (index < 0) {
            index = ListView_GetNextItem(this->listView, -1, LVNI_SELECTED);
        }
        MemItem* fi = this->filter.GetAt(index);
        if (fi == NULL) {
            StatusBar::SetText(this->status, StatusBar::PART_NAME, NULL);
            StatusBar::SetText(this->status, StatusBar::PART_GROUP, NULL);
        } else {
            StatusBar::SetText(this->status, StatusBar::PART_NAME, this->cache.GetItemName(fi->group, fi->item));
            StatusBar::SetText(this->status, StatusBar::PART_GROUP, this->cache.GetGroupName(fi->group));
        }

        HMENU hm = ::GetMenu(hwnd);
        UINT flags = MF_BYCOMMAND | (count > 0 ? MF_ENABLED : MF_GRAYED);
        UINT state = ::EnableMenuItem(hm, IDM_SAVE_PNG, flags);
        if (state != 0xffffffff && ((state ^ flags) & MF_GRAYED) == MF_GRAYED) {
            ::EnableMenuItem(hm, IDM_SAVE_BMP, flags);
            ::EnableMenuItem(hm, IDM_SAVE_ICO, flags);
            ::DrawMenuBar(hwnd);
        }
    }

    void UpdateLayout(HWND hwnd) throw()
    {
        ::SendMessage(this->edit, WM_SIZE, 0, 0);
        ::SendMessage(this->status, WM_SIZE, 0, 0);
        RECT rect;
        ::GetEffectiveClientRect(hwnd, &rect, this->ctlArray);
        ::MoveWindow(this->listView, rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, TRUE);
    }

    int LoadImage(MemItem* fi) throw()
    {
        if (fi->image < 0) {
            fi->image = I_IMAGENONE;
            const ITEM* item = this->cache.GetItem(fi->group, fi->item);
            if (this->bitmap == NULL) {
                this->bitmap = Create32bppTopDownDIB(IMGWIDTH, IMGHEIGHT, &this->bits);
            }
            if (this->bitmap != NULL) {
                HRESULT hr = this->cache.CopyItemImage(item, this->bits, IMGWIDTH * IMGHEIGHT * 4);
                if (SUCCEEDED(hr)) {
                    int i;
                    hr = this->imageList->Add(this->bitmap, NULL, &i);
                    if (SUCCEEDED(hr)) {
                        fi->image = i;
                    }
                }
            }
        }
        return fi->image;
    }

    LPCWSTR BrowseFolderPath(HWND hwnd) throw()
    {
        HRESULT hr;
        LPWSTR path = NULL;
        LPCWSTR out = NULL;
        int i;
        FAIL_HR(hr = BrowseForFolder(hwnd, &path));
        FAIL_HR(hr = this->folders.Add(path, &i));
        out = this->folders.GetAt(i);
Fail:
        ::SHFree(path);
        return out;
    }

    LPCWSTR GetFolderPath(HWND hwnd, int id) throw()
    {
        if (id == IDM_PNG_FOLDER || id == IDM_BMP_FOLDER || id == IDM_ICO_FOLDER) {
            return this->BrowseFolderPath(hwnd);
        }
        LPCWSTR path = NULL;
        path = this->folders.GetAt(id - IDM_PNG_FOLDER - 1);
        if (path == NULL) {
            path = this->folders.GetAt(id - IDM_BMP_FOLDER - 1);
        }
        if (path == NULL) {
            path = this->folders.GetAt(id - IDM_ICO_FOLDER - 1);
        }
        return path;
    }

    static COPYCOMMAND GetCopyCommand(int id) throw()
    {
        if (IDM_PNG_FOLDER <= id && id <= IDM_PNG_FOLDER + FolderList::MAX_ITEMS) {
            return &MainWindow::CopyPNG;
        }
        if (IDM_BMP_FOLDER <= id && id <= IDM_BMP_FOLDER + FolderList::MAX_ITEMS) {
            return &MainWindow::CopyBMP;
        }
        if (IDM_ICO_FOLDER <= id && id <= IDM_ICO_FOLDER + FolderList::MAX_ITEMS) {
            return &MainWindow::CopyICO;
        }
        return NULL;
    }

    HRESULT GetFileName(LPWSTR out, size_t cch, LPCWSTR dir, LPCWSTR group, LPCWSTR name) const throw()
    {
        return ::StringCchPrintfW(out, cch, L"%s\\%s\\%s\\%s.png", this->appPath, dir, group, name);
    }

    HRESULT GetSelectedItems(ItemArray& array) const throw()
    {
        HRESULT hr;
        int index = -1;
        const UINT cSelected = ListView_GetSelectedCount(this->listView);
        FAIL_HR(hr = array.Create() ? S_OK : E_OUTOFMEMORY);
        if (cSelected == 0) {
            return S_FALSE;
        }
        array.Grow(cSelected);

        while ((index = ListView_GetNextItem(this->listView, index, LVNI_SELECTED)) != -1) {
            MemItem* item = this->filter.GetAt(index);
            if (item == NULL) {
                continue;
            }
            FAIL_HR(hr = array.Add(item) != DA_ERR ? S_OK : E_OUTOFMEMORY);
        }
        hr = array.GetCount() >= 0 ? S_OK : S_FALSE;
Fail:
        return hr;
    }

    HRESULT CopyIcons(HWND hwnd, LPCWSTR outDir, COPYCOMMAND command) throw()
    {
        HRESULT hr;
        ItemArray array;

        FAIL_HR(hr = this->GetSelectedItems(array));
        if (hr == S_FALSE) {
            StatusBar::SetText(this->status, StatusBar::PART_STAT, NULL);
            return hr;
        }
        for (int i = 0, count = array.GetCount(); i < count; i++) {
            MemItem* item = array.FastGetAt(i);
            WCHAR src32[MAX_PATH], src16[MAX_PATH];
            const wchar_t* group = this->cache.GetGroupName(item->group);
            const wchar_t* name = this->cache.GetItemName(item->group, item->item);
            FAIL_HR(hr = this->GetFileName(src32, _countof(src32), L"FatCow_Icons32x32_Color", group, name));
            FAIL_HR(hr = this->GetFileName(src16, _countof(src16), L"FatCow_Icons16x16_Color", group, name));
            hr = (this->*command)(hwnd, outDir, name, src32, src16);
            if (FAILED(hr)) {
                ErrorMessage(hwnd, name);
            }
            FAIL_HR(hr);
        }
        StatusBar::SetText(this->status, StatusBar::PART_STAT, L"All done.");
        return hr;

Fail:
        StatusBar::SetText(this->status, StatusBar::PART_STAT, L"Oops.");
        return hr;
    }

    HRESULT CopyPNG(HWND hwnd, LPCWSTR outDir, LPCWSTR name, LPCWSTR src32, LPCWSTR src16) throw()
    {
        HRESULT hr;
        WCHAR dst32[MAX_PATH], dst16[MAX_PATH];
        FAIL_HR(hr = ::StringCchPrintfW(dst32, _countof(dst32), L"%s\\%s.png", outDir, name));
        FAIL_HR(hr = ::StringCchPrintfW(dst16, _countof(dst16), L"%s\\%s_16.png", outDir, name));
        if (CanOverwrite(hwnd, dst32)) {
            FAIL_HR(hr = this->CopyFile(src32, dst32));
        }
        if (CanOverwrite(hwnd, dst16)) {
            FAIL_HR(hr = this->CopyFile(src16, dst16));
        }
Fail:
        return hr;
    }

    HRESULT CopyBMP(HWND hwnd, LPCWSTR outDir, LPCWSTR name, LPCWSTR src32, LPCWSTR src16) throw()
    {
        HRESULT hr;
        WCHAR dst32[MAX_PATH], dst16[MAX_PATH];
        FAIL_HR(hr = ::StringCchPrintfW(dst32, _countof(dst32), L"%s\\L", outDir));
        FAIL_HR(hr = ::StringCchPrintfW(dst16, _countof(dst16), L"%s\\S", outDir));
        FAIL_HR(hr = CreateDirectory(dst32));
        FAIL_HR(hr = CreateDirectory(dst16));
        FAIL_HR(hr = ::StringCchPrintfW(dst32, _countof(dst32), L"%s\\L\\%s.bmp", outDir, name));
        FAIL_HR(hr = ::StringCchPrintfW(dst16, _countof(dst16), L"%s\\S\\%s.bmp", outDir, name));
        if (CanOverwrite(hwnd, dst32)) {
            FAIL_HR(hr = this->SaveAsBitmap(src32, dst32));
            CopyTimeStamps(src32, dst32);
        }
        if (CanOverwrite(hwnd, dst16)) {
            FAIL_HR(hr = this->SaveAsBitmap(src16, dst16));
            CopyTimeStamps(src16, dst16);
        }
Fail:
        return hr;
    }

    HRESULT CopyICO(HWND hwnd, LPCWSTR outDir, LPCWSTR name, LPCWSTR src32, LPCWSTR src16) throw()
    {
        HRESULT hr;
        WCHAR dst[MAX_PATH];
        UINT32 size32, size16, mask32, mask16;
        DibImage image32, image16;
        UINT cx32, cy32, cx16, cy16;
        BYTE* mask = NULL;
        LazyIconDir hdr = {};
        BITMAPINFOHEADER bih = { sizeof(bih), 0, 0, 1, 32, BI_RGB };
        ComPtr<IStream> stream;

        FAIL_HR(hr = ::StringCchPrintfW(dst, _countof(dst), L"%s\\%s.ico", outDir, name));
        FAIL_HR(hr = this->LoadAsBitmap(src32, image32));
        FAIL_HR(hr = this->LoadAsBitmap(src16, image16));
        cx32 = image32.GetWidth();
        cy32 = image32.GetHeight();
        cx16 = image16.GetWidth();
        cy16 = image16.GetHeight();
        FAIL_HR(hr = cx32 <= 256 && cy32 <= 256 ? S_OK : E_FAIL);
        FAIL_HR(hr = cx16 <= 256 && cy16 <= 256 ? S_OK : E_FAIL);
        FAIL_HR(hr = SafeAlloc(GetStride(__max(cx16, cx32), 1) * __max(cy16, cy32), mask));

        size32 = GetStride(cx32, 32) * cy32;
        size16 = GetStride(cx16, 32) * cy16;
        mask32 = GetStride(cx32, 1) * cy32;
        mask16 = GetStride(cx16, 1) * cy16;
        hdr.one = 1;
        hdr.two = 2;
        hdr.entries[0].width = cx32 & 0xff;
        hdr.entries[0].height = cy32 & 0xff;
        hdr.entries[0].one = 1;
        hdr.entries[0].thirtytwo = 32;
        hdr.entries[0].size = sizeof(bih) + size32 + mask32;
        hdr.entries[0].offset = sizeof(LazyIconDir);
        hdr.entries[1].width = cx16 & 0xff;
        hdr.entries[1].height = cy16 & 0xff;
        hdr.entries[1].one = 1;
        hdr.entries[1].thirtytwo = 32;
        hdr.entries[1].size = sizeof(bih) + size16 + mask16;
        hdr.entries[1].offset = sizeof(LazyIconDir) + sizeof(bih) + size32 + mask32;

        if (CanOverwrite(hwnd, dst)) {
            FAIL_HR(hr = ::SHCreateStreamOnFileEx(dst, STGM_CREATE | STGM_WRITE | STGM_SHARE_EXCLUSIVE, 0, TRUE, NULL, &stream));
            FAIL_HR(hr = StreamWrite(stream, &hdr, sizeof(hdr)));
            bih.biWidth = cx32;
            bih.biHeight = cy32 * 2;
            FAIL_HR(hr = StreamWrite(stream, &bih, sizeof(bih)));
            FAIL_HR(hr = StreamWrite(stream, image32.GetBits(), size32));
            ZeroMemory(mask, mask32);
            image32.CopyMask(mask);
            FAIL_HR(hr = StreamWrite(stream, mask, mask32));
            bih.biWidth = cx16;
            bih.biHeight = cy16 * 2;
            FAIL_HR(hr = StreamWrite(stream, &bih, sizeof(bih)));
            FAIL_HR(hr = StreamWrite(stream, image16.GetBits(), size16));
            ZeroMemory(mask, mask16);
            image16.CopyMask(mask);
            FAIL_HR(hr = StreamWrite(stream, mask, mask16));

            stream.Release();
            CopyTimeStamps(src32, dst);
        }

Fail:
        SafeFree(mask);
        return hr;
    }

    HRESULT LoadAsBitmap(LPCWSTR file, __inout DibImage& image) throw()
    {
        HRESULT hr;
        ComPtr<IWICImagingFactory> factory;
        FAIL_HR(hr = this->GetWICFactory(&factory));
        FAIL_HR(hr = image.Init(factory, file));
        image.FlipVertical();

Fail:
        return hr;
    }

    HRESULT SaveAsBitmap(LPCWSTR src, LPCWSTR dst) throw()
    {
        HRESULT hr;
        ComPtr<IStream> stream;
        DibImage image;

        FAIL_HR(hr = this->LoadAsBitmap(src, image));
        FAIL_HR(hr = ::SHCreateStreamOnFileEx(dst, STGM_CREATE | STGM_WRITE | STGM_SHARE_EXCLUSIVE, 0, TRUE, NULL, &stream));
        {
            const UINT cx = image.GetWidth();
            const UINT cy = image.GetHeight();
            const UINT32 size = GetStride(cx, 32) * cy;
            BITMAPINFOHEADER bih = { sizeof(bih), static_cast<LONG>(cx), static_cast<LONG>(cy), 1, 32, BI_RGB };
            BITMAPFILEHEADER bfh = { 'MB', sizeof(bfh) + sizeof(bih) + size, 0, 0, sizeof(bfh) + sizeof(bih) };
            FAIL_HR(hr = StreamWrite(stream, &bfh, sizeof(bfh)));
            FAIL_HR(hr = StreamWrite(stream, &bih, sizeof(bih)));
            FAIL_HR(hr = StreamWrite(stream, image.GetBits(), size));
        }

Fail:
        return hr;
    }

    static HRESULT CreateDirectory(LPCWSTR path) throw()
    {
        BOOL ret = ::CreateDirectoryW(path, NULL);
        if (ret || ::GetLastError() == ERROR_ALREADY_EXISTS) {
            return S_OK;
        }
        return E_FAIL;
    }

    static HRESULT CopyFile(LPCWSTR src, LPCWSTR dst) throw()
    {
        if (::CopyFileW(src, dst, TRUE)) {
            return S_OK;
        }
        return E_FAIL;
    }

    static HRESULT CopyTimeStamps(LPCWSTR src, LPCWSTR dst) throw()
    {
        HRESULT hr;
        FILE_BASIC_INFO fbiDst = {};
        FILE_BASIC_INFO fbiSrc = {};

        HANDLE hfSrc = ::CreateFileW(src, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
        HANDLE hfDst = ::CreateFileW(dst, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);

        FAIL_HR(hr = hfSrc != INVALID_HANDLE_VALUE ? S_OK : E_FAIL);
        FAIL_HR(hr = hfDst != INVALID_HANDLE_VALUE ? S_OK : E_FAIL);
        FAIL_HR(hr = ::GetFileInformationByHandleEx(hfSrc, FileBasicInfo, &fbiSrc, sizeof(fbiSrc)) ? S_OK : E_FAIL);
        FAIL_HR(hr = ::GetFileInformationByHandleEx(hfDst, FileBasicInfo, &fbiDst, sizeof(fbiDst)) ? S_OK : E_FAIL);

        fbiDst.CreationTime     = fbiSrc.CreationTime;
        fbiDst.LastAccessTime   = fbiSrc.LastAccessTime;
        fbiDst.LastWriteTime    = fbiSrc.LastWriteTime;
        fbiDst.ChangeTime       = fbiSrc.ChangeTime;
        FAIL_HR(hr = ::SetFileInformationByHandle(hfDst, FileBasicInfo, &fbiDst, sizeof(fbiDst)) ? S_OK : E_FAIL);

Fail:
        if (hfSrc != INVALID_HANDLE_VALUE) {
            ::CloseHandle(hfSrc);
        }
        if (hfDst != INVALID_HANDLE_VALUE) {
            ::CloseHandle(hfDst);
        }
        return hr;
    }

    static bool CanOverwrite(HWND hwnd, LPCWSTR path) throw()
    {
        if (::PathFileExistsW(path) == FALSE) {
            return true;
        }
        WCHAR message[MAX_PATH + 64];
        ::StringCchCopyW(message, _countof(message), L"Do you want to overwrite:\n");
        ::StringCchCatW(message, _countof(message), path);
        if (::MessageBoxW(hwnd, message, L"Overwrite", MB_YESNO) == IDYES) {
            return true;
        }
        return false;
    }

    static void ErrorMessage(HWND hwnd, LPCWSTR name) throw()
    {
        WCHAR message[MAX_PATH + 64];
        ::StringCchCopyW(message, _countof(message), L"Could not copy:\n");
        ::StringCchCatW(message, _countof(message), name);
        ::MessageBoxW(hwnd, message, L"Sorry", MB_OK);
    }

    LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) throw()
    {
        switch (message) {
        HANDLE_MSG(hwnd, WM_CREATE, OnCreate);
        HANDLE_MSG(hwnd, WM_DESTROY, OnDestroy);
        HANDLE_MSG(hwnd, WM_COMMAND, OnCommand);
        HANDLE_MSG(hwnd, WM_WINDOWPOSCHANGED, OnWindowPosChanged);
        HANDLE_MSG(hwnd, WM_SETTINGCHANGE, OnSettingChange);
        HANDLE_MSG(hwnd, WM_ACTIVATE, OnActivate);
        HANDLE_MSG(hwnd, WM_NOTIFY, OnNotify);
        HANDLE_MSG(hwnd, WM_INITMENU, OnInitMenu);
        HANDLE_MSG(hwnd, WM_INITMENUPOPUP, OnInitMenuPopup);
        HANDLE_MSG(hwnd, WM_CONTEXTMENU, OnContextMenu);
        }
        return ::DefWindowProc(hwnd, message, wParam, lParam);
    }

    bool OnCreate(HWND hwnd, LPCREATESTRUCT lpcs) throw()
    {
        if (lpcs == NULL) {
            return false;
        }
        this->edit = this->search.Create(hwnd, IDC_EDIT);
        if (this->edit == NULL) {
            return false;
        }
        this->listView = ::CreateWindowEx(0, WC_LISTVIEW, NULL,
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
            | LVS_SHOWSELALWAYS | LVS_AUTOARRANGE | LVS_ALIGNTOP | LVS_SHAREIMAGELISTS | LVS_OWNERDATA,
            0, 0, 0, 0,
            hwnd,
            reinterpret_cast<HMENU>(IntToPtr(IDC_LISTVIEW)),
            HINST_THISCOMPONENT, NULL);
        if (this->listView == NULL) {
            return false;
        }
        this->status = StatusBar::Create(hwnd, IDC_STATUSBAR);
        if (this->status == NULL) {
            return false;
        }
        HMENU hm = ::GetMenu(hwnd);
        this->popup = GetSubMenuById(hm, IDM_POPUP);
        ::RemoveMenu(hm, IDM_POPUP, MF_BYCOMMAND);

        this->ctlArray[0] = -1;
        this->ctlArray[1] = PtrToInt(hm);
        this->ctlArray[2] = -1;
        this->ctlArray[3] = IDC_EDIT;
        this->ctlArray[4] = IDM_STATUSBAR;
        this->ctlArray[5] = IDC_STATUSBAR;
        this->ctlArray[6] = 0;
        this->ctlArray[7] = 0;

        ::SetWindowTheme(this->listView, L"Explorer", NULL);
        ListView_SetExtendedListViewStyle(this->listView, LVS_EX_CHECKBOXES | LVS_EX_INFOTIP | LVS_EX_UNDERLINEHOT | LVS_EX_DOUBLEBUFFER | LVS_EX_AUTOCHECKSELECT);

        this->pngMenu.Init(hm, IDM_SAVE_PNG, IDM_PNG_FOLDER);
        this->bmpMenu.Init(hm, IDM_SAVE_BMP, IDM_BMP_FOLDER);
        this->icoMenu.Init(hm, IDM_SAVE_ICO, IDM_ICO_FOLDER);

        SetSubMenuById(this->popup, IDM_SAVE_PNG, this->pngMenu.GetMenu());
        SetSubMenuById(this->popup, IDM_SAVE_BMP, this->bmpMenu.GetMenu());
        SetSubMenuById(this->popup, IDM_SAVE_ICO, this->icoMenu.GetMenu());

        this->UpdateLayout(hwnd);
        return true;
    }

    void OnDestroy(HWND) throw()
    {
        this->folders.Save(this->listPath);
        ::PostQuitMessage(0);
    }

    void OnCommand(HWND hwnd, int id, HWND, UINT code) throw()
    {
        switch (id) {
        case IDC_EDIT:
            this->OnCommandEdit(hwnd, id, code);
            break;
        case IDM_COPY_NAME:
        case IDA_COPY_NAME:
            this->OnCommandCopyName(hwnd, id);
            break;
        case IDM_EXIT_APP:
            this->OnCommandExitApp(hwnd, id);
            break;
        case IDM_STATUSBAR:
            this->OnCommandStatusBar(hwnd, id);
            break;
        case IDA_FOCUS_EDIT:
            this->OnCommandFocusEdit(hwnd, id);
            break;
        case IDA_PREV_CONTROL:
        case IDA_NEXT_CONTROL:
            this->OnCommandFocusControl(hwnd, id);
            break;
        default:
            if (id >= IDM_PNG_FOLDER || id >= IDM_BMP_FOLDER || id >= IDM_ICO_FOLDER) {
                this->OnCommandSave(hwnd, id);
            }
            break;
        }
    }

    void OnCommandEdit(HWND hwnd, int, UINT code) throw()
    {
        if (code == EN_CHANGE) {
            this->PopulateItems(hwnd);
        }
    }

    void OnCommandCopyName(HWND hwnd, int id) throw()
    {
        if (id == IDA_COPY_NAME && ::GetFocus() != this->listView) {
            return;
        }

        ItemArray array;
        if (this->GetSelectedItems(array) != S_OK) {
            return;
        }

        size_t cch = 1;
        for (int i = 0, count = array.GetCount(); i < count; i++) {
            MemItem* item = array.FastGetAt(i);
            const wchar_t* name = this->cache.GetItemName(item->group, item->item);
            cch += ::wcslen(name) + 2; // sizeof(CRLF)
        }

        const size_t cb = cch * sizeof(wchar_t);
        HGLOBAL hglob = ::GlobalAlloc(GHND | GMEM_SHARE, cb);
        if (hglob != NULL) {
            wchar_t* dst = reinterpret_cast<wchar_t*>(::GlobalLock(hglob));
            for (int i = 0, count = array.GetCount(); i < count; i++) {
                if (i > 0) {
                    *dst++ = L'\r';
                    *dst++ = L'\n';
                }
                MemItem* item = array.FastGetAt(i);
                const wchar_t* src = this->cache.GetItemName(item->group, item->item);
                while (*src != L'\0') {
                    *dst++ = *src++;
                }
            }
            ::GlobalUnlock(hglob);
        }
        if (hglob != NULL) {
            HANDLE handle = NULL;
            if (::OpenClipboard(hwnd)) {
                ::EmptyClipboard();
                handle = ::SetClipboardData(CF_UNICODETEXT, hglob);
                ::CloseClipboard();
            }
            if (handle == NULL) {
                ::GlobalFree(hglob);
            }
        }
    }

    void OnCommandSave(HWND hwnd, int id) throw()
    {
        const UINT cSelected = ListView_GetSelectedCount(this->listView);
        if (cSelected == 0) {
            return;
        }
        const wchar_t* outPath = this->GetFolderPath(hwnd, id);
        COPYCOMMAND command = GetCopyCommand(id);
        if (outPath != NULL && command != NULL) {
            this->CopyIcons(hwnd, outPath, command);
        }
    }

    void OnCommandExitApp(HWND hwnd, int) throw()
    {
        ::SendMessage(hwnd, WM_CLOSE, 0, 0);
    }

    void OnCommandStatusBar(HWND hwnd, int id) throw()
    {
        ::ShowHideMenuCtl(hwnd, id, this->ctlArray);
        this->UpdateLayout(hwnd);
    }

    void OnCommandFocusEdit(HWND, int) throw()
    {
        Edit_SetSel(this->edit, 0, -1);
        ::SetFocus(this->edit);
    }

    void OnCommandFocusControl(HWND, int) throw()
    {
        HWND hwnd = ::GetFocus();
        if (hwnd == this->edit) {
            hwnd = this->listView;
        } else {
            hwnd = this->edit;
        }
        ::SetFocus(hwnd);
    }

    void OnWindowPosChanged(HWND hwnd, const WINDOWPOS* pwp) throw()
    {
        if (pwp == NULL) {
            return;
        }
        if (pwp->flags & SWP_SHOWWINDOW) {
            if (this->shown == false) {
                this->shown = true;
                this->LoadIcons(hwnd);
            }
        }
        if ((pwp->flags & SWP_NOSIZE) == 0) {
            this->UpdateLayout(hwnd);
        }
    }

    void OnSettingChange(HWND hwnd, UINT, LPCTSTR) throw()
    {
        this->UpdateLayout(hwnd);
    }

    void OnActivate(HWND, UINT state, HWND, BOOL) throw()
    {
        if (state != WA_INACTIVE) {
            ::SendMessage(this->listView, WM_CHANGEUISTATE, UIS_INITIALIZE, 0);
            ::SetFocus(this->edit);
        }
    }

    LRESULT OnNotify(HWND hwnd, int idFrom, NMHDR* pnmhdr) throw()
    {
        if (idFrom != IDC_LISTVIEW || pnmhdr == NULL) {
            return 0;
        }

        switch (pnmhdr->code) {
        case LVN_ITEMCHANGED:
            return this->OnNotifyListViewItemChanged(hwnd, CONTAINING_RECORD(pnmhdr, NMLISTVIEW, hdr));
        case LVN_GETDISPINFOW:
            return this->OnNotifyListViewGetDispInfo(hwnd, CONTAINING_RECORD(pnmhdr, NMLVDISPINFOW, hdr));
        }
        return 0;
    }

    LRESULT OnNotifyListViewItemChanged(HWND hwnd, NMLISTVIEW* lplv) throw()
    {
        if ((lplv->uChanged & LVIF_STATE) != 0) {
            if (((lplv->uNewState ^ lplv->uOldState) & LVIS_SELECTED) == LVIS_SELECTED) {
                this->UpdateStatus(hwnd);
            }
        }
        return 0;
    }

    LRESULT OnNotifyListViewGetDispInfo(HWND, NMLVDISPINFOW* lplvdi) throw()
    {
        MemItem* fi = this->filter.GetAt(lplvdi->item.iItem);
        if (fi == NULL) {
            return 0;
        }
        if (lplvdi->item.mask & LVIF_TEXT) {
            ::StringCchCopyW(lplvdi->item.pszText, lplvdi->item.cchTextMax, this->cache.GetItemName(fi->group, fi->item));
        }
        if (lplvdi->item.mask & LVIF_IMAGE) {
            lplvdi->item.iImage = this->LoadImage(fi);
        }
        lplvdi->item.mask |= LVIF_DI_SETITEM;
        return 0;
    }

    void OnInitMenu(HWND, HMENU hMenu) throw()
    {
        UINT count = ListView_GetSelectedCount(this->listView);
        UINT flags = MF_BYCOMMAND | (count > 0 ? MF_ENABLED : MF_GRAYED);
        ::EnableMenuItem(hMenu, IDM_COPY_NAME, flags);
        ::EnableMenuItem(hMenu, IDM_SAVE_PNG, flags);
        ::EnableMenuItem(hMenu, IDM_SAVE_BMP, flags);
        ::EnableMenuItem(hMenu, IDM_SAVE_ICO, flags);
    }

    void OnInitMenuPopup(HWND, HMENU hMenu, UINT, BOOL) throw()
    {
        this->pngMenu.HandleInitMenuPopup(hMenu, this->folders);
        this->bmpMenu.HandleInitMenuPopup(hMenu, this->folders);
        this->icoMenu.HandleInitMenuPopup(hMenu, this->folders);
    }

    void OnContextMenu(HWND hwnd, HWND hwndContext, int xPos, int yPos) throw()
    {
        if (hwndContext != this->listView) {
            FORWARD_WM_CONTEXTMENU(hwnd, hwndContext, xPos, yPos, ::DefWindowProc);
            return;
        }
        int iItem = ListView_GetNextItem(this->listView, -1, LVNI_SELECTED | LVNI_FOCUSED);
        if (iItem < 0) {
            return;
        }
        POINT pt = { xPos, yPos };
        if (MAKELONG(pt.x, pt.y) == -1) {
            ListView_EnsureVisible(this->listView, iItem, FALSE);
            RECT rcItem = {};
            ListView_GetItemRect(this->listView, iItem, &rcItem, LVIR_ICON);
            pt.x = rcItem.left + (rcItem.right - rcItem.left) / 2;
            pt.y = rcItem.top  + (rcItem.bottom - rcItem.top) / 2;
            ::ClientToScreen(this->listView, &pt);
        }
        ::TrackPopupMenuEx(popup, TPM_RIGHTBUTTON | TPM_LEFTBUTTON, pt.x, pt.y, hwnd, NULL);
    }

public:
    MainWindow() throw()
        : edit()
        , listView()
        , status()
        , popup()
        , bitmap()
        , bits()
        , shown()
        , items()
        , cItems()
    {
    }

    ~MainWindow() throw()
    {
        if (this->popup != NULL) {
            ::RemoveMenu(this->popup, IDM_SAVE_PNG, MF_BYCOMMAND);
            ::RemoveMenu(this->popup, IDM_SAVE_BMP, MF_BYCOMMAND);
            ::RemoveMenu(this->popup, IDM_SAVE_ICO, MF_BYCOMMAND);
            ::DestroyMenu(this->popup);
        }
        if (this->bitmap != NULL) {
            DeleteBitmap(this->bitmap);
        }
        delete[] items;
    }

    HWND Create() throw()
    {
        ::InitCommonControls();
        LPCTSTR menu = MAKEINTRESOURCE(IDR_MAIN);
        const ATOM atom = __super::Register(MAKEINTATOM(911), menu);
        if (atom == 0) {
            return NULL;
        }
        return __super::Create(atom, L"Web Icon Browser");
    }
};


class CCoInitialize
{
public:
    HRESULT m_hr;

    CCoInitialize() throw()
        : m_hr(CoInitialize(NULL))
    {
    }

    ~CCoInitialize() throw()
    {
        if (SUCCEEDED(m_hr)) {
            CoUninitialize();
        }
    }

    HRESULT GetResult() const throw()
    {
        return m_hr;
    }
};

class CAllocConsole
{
public:
    CAllocConsole() throw()
    {
        ::AllocConsole();
    }
    ~CAllocConsole() throw()
    {
        ::FreeConsole();
    }
};

int WINAPI wWinMain(HINSTANCE, HINSTANCE, LPWSTR, int) throw()
{
#ifdef USETRACE
    CAllocConsole con;
#endif
    CCoInitialize init;
    MainWindow wnd;
    HWND hwnd = wnd.Create();
    if (hwnd == NULL) {
        return EXIT_FAILURE;
    }
    MSG msg;
    msg.wParam = EXIT_FAILURE;
    HACCEL hacc = ::LoadAccelerators(HINST_THISCOMPONENT, MAKEINTRESOURCE(IDR_MAIN));
    while (::GetMessage(&msg, NULL, 0, 0)) {
        if (hacc != NULL && (hwnd == msg.hwnd || ::IsChild(hwnd, msg.hwnd))) {
            if (::TranslateAccelerator(hwnd, hacc, &msg)) {
                continue;
            }
        }
        ::DispatchMessage(&msg);
        ::TranslateMessage(&msg);
    }
    return static_cast<int>(msg.wParam);
}
