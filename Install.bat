echo Install Tools ...

cd setuptools-18.7
python setup.py install

cd ..
cd xlrd-0.9.4
python setup.py install

cd ..
cd protobuf-V3.0/python
python setup.py install

cd ..
echo Modify Env Path
set PATH = %PATH%;%~dp0protobuf-V3.0/src/


echo Install Ok

pause
