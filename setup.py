from distutils.core import setup
import py2exe


setup(
   windows=['ExlGeo_v3.pyw'],
   options={"py2exe":{
                "skip_archive": True
                    }

   }
)
#C:\Python27\python.exe setup.py py2exe