from distutils.core import setup
import py2exe




setup(
   console=['ExlGeo_v3.pyw'],
   options={
       "py2exe":{
            "skip_archive": True
                }
        }
)

#python setup.py py2exe