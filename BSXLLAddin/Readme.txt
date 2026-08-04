"BSXLLAddin" - BLUESOFTS XLL ADD-IN BUILDER
Version: 1.0.0 - Updated 2026-07-30

(English)
INTRODUCTION TO "BSXLLAddin"

"BSXLLAddin" is a library that supports creating XLL add-ins for Excel in Delphi or C++ Builder in the easiest way.
"BSXLLAddin" allows you to create UDF functions in two ways:

+ Method 1: Using basic Delphi language
With this method, you need to use basic Delphi language and the data types provided by Delphi to create functions in Delphi style. For example:

function MySum(const v1, v2: OleVariant): OleVariant; stdcall;
begin
	VariantInit(Result);
	Result := v1 + v2;
end;

(*) This method requires the following:
	- The function must return an OleVariant type and have a "stdcall" declaration at the end.
	- Parameters must be declared as "const" and their type must be OleVariant.

+ Method 2: Creating functions using the "C API in Excel" style
This method is much more complex; you need good knowledge of memory management and pointer types.

HOW TO CREATE XLL USING "Bluesofts XLL Add-in"
+ Go to the File menu -> New -> "Other..." -> "Bluesofts XLL Add-in"
+ Declare your XLL add-in information. Click "Create".
The system will create a sample project with example functions for you to understand and follow.

If you need the PRO package with full features and author support during the XLL add-in creation process, contact:
Author: Nguyen Duy Tuan - tuanktcdcn@yahoo.com - Tel: (+84) 904210337.

//-----------(Ngôn gữ Việt Nam)---------------//
GIỚI THIỆU "BSXLLAddin"
"BSXLLAddin" là thư viện hỗ trợ tạo XLL add-in cho Excel trong Delphi hoặc C++ Builder theo cách dễ dàng nhất.
"BSXLLAddin" cho phép bạn tạo các hàm UDF theo hai cách:
+ Cách 1: Dùng ngôn ngữ Delphi cơ bản
Cách này bạn cần dùng ngôn ngữ Delphi cơ bản, kiểu dữ liệu Delphi đã cung cấp để tạo ra các hàm theo kiểu Delphi là được. Ví dụ

function MySum(const v1, v2: OleVariant): OleVariant; stdcall;
begin
	VariantInit(Result);
	Result := v1 + v2;
end;

(*) Theo cách này yêu cầu như sau:
	- Hàm trả về kiểu OleVariant và có khai báo "stdcall" ở cuối. 
	- Các tham số phải khai báo "const" và kiểu phải là OleVariant.

+ Cách 2: Tạo hàm theo kiểu "C API in Excel"
Với cách này phức tạp hơn rất nhiều, bạn cần có kiến thức tốt về quản bộ nhớ và kiểu con trỏ - pointer.

CÁCH TẠO XLL BẰNG "BSXLLAddin"
+ Vào menu File -> New -> "Other..." -> "Bluesofts XLL Add-in"
+ Khai báo thông tin xll add-in của bạn. Bấm "Create".
Hệ thống tạo cho bạn dữ án mẫu với các hàm ví dụ để bạn hiểu và làm theo.

Nếu cần gói PRO với full tính năng, được tác giả hỗ trợ trong quá trình tạo xll add-in liên hệ:
Tác giả: Nguyễn Duy Tuân - tuanktcdcn@yahoo.com - Tel: (+84) 904210337.

