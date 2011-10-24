path=C:\Program Files (x86)\Firebird\Firebird_2_5\bin

del temp.foc
del focus.ant
copy focus.foc temp.foc
ren focus.foc focus.Ant

set isc_user=SYSDBA
set isc_password=masterkey

gfix -m -i temp.foc
gfix -m -i temp.foc
gfix -m -i temp.foc

gbak -g -b -z -v -i temp.foc back.gbk
gbak -g -r -z -v back.gbk focus.foc

del back.gbk
del temp.foc

cd \
echo *****  fim do reparar.bat ****
ren
