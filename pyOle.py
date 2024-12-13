import pythoncom
import win32com.server.register
from win32com.server.exception import COMException
import ctypes  

class Controller:
   
   _public_methods_ = ['setValue', 'getValue','curl','extract'] 
   _public_attrs_ = ['value']
   _reg_progid_ = "MyPythonCOM.MySimpleCOMObject"
   _reg_verprogid_ = "MyPythonCOM.MySimpleCOMObject.1"
   _reg_desc_ = "Python Test COM Server"
   _reg_class_spec_ = "pyOle.Controller"   
   _reg_clsid_ = "{5b0a1053-a254-4981-b32f-ec172b93c0ea}"

   def __init__(self):
      self.value = None

   def setValue(self, new_value):
      try:
         self.value = new_value
         return
      except Exception as e:
         raise COMException(f"Error setting value: {str(e)}")

   def getValue(self):
      try:
         return self.value if self.value is not None else "Value not set"
      except Exception as e:
         raise COMException(f"Error getting value: {str(e)}")
      
   def extract(self, pdf):
      try:
         import PyPDF2
         from PyPDF2 import PdfReader

         reader = PdfReader(pdf)

         text = ""
         for page in reader.pages:
            text += page.extract_text()
         return text
      
      except Exception as e:
         raise COMException(f"Error curl: {str(e)}")      

def show_message(message, title="Information"):
   ctypes.windll.user32.MessageBoxW(0, message, title, 0x40)

def DllRegisterServer(silent=False):
   try:
      pythoncom.CoInitialize()
      win32com.server.register.UseCommandLine(Controller)
      message = "Object registered successfully."
   except Exception as e:
      message = f"Error registering object: {e}"
   finally:
      pythoncom.CoUninitialize()
   if not silent:
      show_message(message, "Object registered")

def DllUnregisterServer(silent=False):
   try:
      pythoncom.CoInitialize()
      win32com.server.register.UnregisterServer(
         Controller._reg_clsid_,
         Controller._reg_progid_
      )
      message = "Object unregistered successfully."
   except Exception as e:
      message = f"Error unregistering object: {e}"
   finally:
      pythoncom.CoUninitialize()
   if not silent:
      show_message(message, "Object unregistered")

if __name__ == '__main__':
   import sys
   silent = False
   if len(sys.argv) > 1:
      if len(sys.argv) > 2:   
         if sys.argv[2] == '--silent':
            silent = True
      if sys.argv[1] == '--register':
         DllRegisterServer(silent)
      elif sys.argv[1] == '--unregister':
         DllUnregisterServer(silent)
      elif "/Automate" in sys.argv:
         import win32com.server.localserver
         win32com.server.localserver.serve([Controller._reg_clsid_])
   else:
         show_message("Use:\npyOle.exe --register\npyOle.exe --unregister\n--silent: Silent mode", "Help")