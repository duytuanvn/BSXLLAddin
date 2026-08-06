# BSXLLAddin - BLUESOFTS XLL ADD-IN BUILDER
Version: 1.0.0 - Updated 2026-07-30

**INTRODUCTION TO "BSXLLAddin"**

"BSXLLAddin" is a library that supports creating XLL add-ins for Excel in Delphi in the easiest way.
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
    //fn := AFunctions.Add('MySum', @MySum, True); //The third parameter is set to True, causing this UDF to run asynchronously. While the function is executing, #CALC! appears on the worksheet.
	fn.Function_help := 'The function returns the sum of a + b.';
		//Add parameters to the function.
	fn.Params.Add('A', 'Enter a number > 0');
	fn.Params.Add('B', 'Enter numbers > 0 and < 100');
	//-----------------
end;
```
<p align="center">
<img width="100%" height="584" alt="BSXLLAddin_MySum" src="https://github.com/user-attachments/assets/d6957954-79a6-4d64-9728-4c7f8064aca1" />
</p>

(*) This method requires the following:
	- The function must return an OleVariant type and have a "stdcall" declaration at the end.
	- Parameters must be declared as "const" and their type must be OleVariant.

+ Method 2: Creating functions using the "C API in Excel" style
This method is much more complex; you need good knowledge of memory management and pointer types.

**HOW TO CREATE XLL USING "Bluesofts XLL Add-in" in Delphi**
+ Run Delphi (You can use the CE version) (Rad Studio)
+ Go to the File menu -> New -> "Other..." -> "Bluesofts XLL Add-in"
+ Declare your XLL add-in information. Click "Create".
<p align="center">
<img width="90%" height="512" alt="BSXLLAddin02" src="https://github.com/user-attachments/assets/e50658d1-eaca-4fa7-9891-b002a55a34f0" />
</p>
<p align="center">
<img width="492" height="358" alt="BSXLLAddin03" src="https://github.com/user-attachments/assets/0ea5167c-983d-4fbe-8469-f91095c7dd5c" />
</p>

The system will create a sample project with example functions for you to understand and follow.
<p align="center">
<img width="90%" height="726" alt="BSXLLAddin06" src="https://github.com/user-attachments/assets/04b6c27d-1230-4e7d-b0f9-a0886a5712b6" />

</p>

It allows the creation of functions and macros with all the features that Microsoft has published in its "C API in Excel".

**Function Types:**
+ ftVolatile //Automatically re-run when the spreadsheet updates.
+ ftMacroSheet //Function with the macro attribute
+ ftFunctionSheet //Standard UDF running within a formula
+ ftThreadSafe //From Excel 2007: Allows execution within a thread.
+ ftClusterSafe //From Excel 2010. Upgrade later

**Macro Type**
+ mtInvisible
+ mtFunction //Your function is a UDF called within a formula on the worksheet.
+ mtCommand //Your function is a macro that runs in "Sheet Macro4.0".

This library updates the latest APIs from Microsoft's "C API in Excel".

**Benefits of using BSXLLAddin:**

Regardless of your Delphi programming proficiency, you can easily create XLLs. The generated functions possess the full range of powerful attributes defined in Microsoft's "C API in Excel" (such as asynchronous execution, multithreading, etc.). The code is concise and convenient! There is no need for deep technical research—that is a task for senior specialists. You can dedicate 100% of your time to your business logic and UDFs!

If you need the PRO package with full features and author support during the XLL add-in creation process, contact:
Author: Nguyen Duy Tuan - tuanktcdcn@yahoo.com - Tel: (+84) 904210337.

### Video help:

<p align="center">
  <a href="https://www.youtube.com/watch?v=BTbwxbC3huo" target="_blank">
    <img src="https://img.youtube.com/vi/BTbwxbC3huo/hqdefault.jpg" alt="Video Tutorial" width="70%">
    <br>
    <img src="https://upload.wikimedia.org/wikipedia/commons/b/b8/YouTube_Logo_2017.svg" alt="YouTube" width="120" style="margin-top: 10px;">
  </a>
  <br>
</p>
