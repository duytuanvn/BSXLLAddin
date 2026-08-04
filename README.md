# BSXLLAddin - BLUESOFTS XLL ADD-IN BUILDER
Version: 1.0.0 - Updated 2026-07-30

**INTRODUCTION TO "BSXLLAddin"**

"BSXLLAddin" is a library that supports creating XLL add-ins for Excel in Delphi or C++ Builder in the easiest way.
"BSXLLAddin" allows you to create UDF functions in two ways:

+ Method 1: Using basic Delphi language
With this method, you need to use basic Delphi language and the data types provided by Delphi to create functions in Delphi style. For example:
```pascal
function MySum(const v1, v2: OleVariant): OleVariant; stdcall;
begin
	VariantInit(Result);
	Result := v1 + v2;
end;

procedure TBSMyXLL.OnAddFunctions(AFunctions: TBSFunctions);
var
	fn: TBSFunction;
begin
	//3. Add UDF function MySum()
	fn := AFunctions.Add('MySum', @MySum);
	fn.Function_help := 'The function returns the sum of a + b.';
		//Add parameters to the function.
	fn.Params.Add('A', 'Enter a number > 0');
	fn.Params.Add('B', 'Enter numbers > 0 and < 100');
	//-----------------
end;
```
<p align="center">
<img width="90%" height="584" alt="BSXLLAddin_MySum" src="https://github.com/user-attachments/assets/d6957954-79a6-4d64-9728-4c7f8064aca1" />
</p>

(*) This method requires the following:
	- The function must return an OleVariant type and have a "stdcall" declaration at the end.
	- Parameters must be declared as "const" and their type must be OleVariant.

+ Method 2: Creating functions using the "C API in Excel" style
This method is much more complex; you need good knowledge of memory management and pointer types.

**HOW TO CREATE XLL USING "Bluesofts XLL Add-in"**

+ Go to the File menu -> New -> "Other..." -> "Bluesofts XLL Add-in"
+ Declare your XLL add-in information. Click "Create".
<p align="center">
<img width="80%" height="512" alt="BSXLLAddin02" src="https://github.com/user-attachments/assets/e50658d1-eaca-4fa7-9891-b002a55a34f0" />


<img width="492" height="358" alt="BSXLLAddin03" src="https://github.com/user-attachments/assets/0ea5167c-983d-4fbe-8469-f91095c7dd5c" />
</p>

The system will create a sample project with example functions for you to understand and follow.
<p align="center">
<img width="90%" height="726" alt="BSXLLAddin06" src="https://github.com/user-attachments/assets/04b6c27d-1230-4e7d-b0f9-a0886a5712b6" />

</p>

It allows the creation of functions and macros with all the features that Microsoft has published in its "C API in Excel".
**Function Types:**
+ ftVolatile
+ ftMacroSheet
+ ftFunctionSheet
+ ftThreadSafe //From Excel 2007
+ ftClusterSafe //From Excel 2010. Upgrade later

**Macro Type**
+ mtInvisible
+ mtFunction
+ mtCommand

This library updates the latest APIs from Microsoft's "C API in Excel".
  
If you need the PRO package with full features and author support during the XLL add-in creation process, contact:
Author: Nguyen Duy Tuan - tuanktcdcn@yahoo.com - Tel: (+84) 904210337.


