FUNCTION Main()

   LOCAL oPy

   IF Empty( oPy := win_oleCreateObject( "MyPythonCOM.MySimpleCOMObject" ) )
      ? "Unable to create oPy COM object"
      RETURN
   ENDIF

   oPy:value = 1
   ? oPy:GetValue() + 1
   ? oPy:extract( hb_cwd() + "sample.pdf") 
   
RETURN
