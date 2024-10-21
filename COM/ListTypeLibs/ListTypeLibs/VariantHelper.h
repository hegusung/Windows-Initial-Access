// See
// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/automat/htm/chap6_7zdz.asp
#include <ObjBase.h>

#include <string>

std::string ws2s(const std::wstring& wstr);
std::wstring s2ws(const std::string& str);

std::string VariantTypeAsString(const VARIANT& v);
std::string        TypeAsString(VARTYPE);

class variant {
public:
    variant(VARIANT);
    ~variant();

    std::string ValueAsString();

    void ChangeType(VARTYPE);

private:
    VARIANT v_;
};