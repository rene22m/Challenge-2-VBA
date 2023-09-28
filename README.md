# Challenge-2-VBA

This are my files regarding the Module's 2 challenge: VBA.

You'll find that in this code I delcared "TotalVolume" and "init" as variant.
While I could have declared them as Double, Excel for Mac gives an overflow error. You can find more about it here: https://techcommunity.microsoft.com/t5/excel/runtime-error-6-overflow-with-dim-double-macos-catalina-excel/m-p/786433
The easier workaround this is to replace Double with Variant.

Also, in order to have the code loop through every worksheet I used this code for reference: https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
