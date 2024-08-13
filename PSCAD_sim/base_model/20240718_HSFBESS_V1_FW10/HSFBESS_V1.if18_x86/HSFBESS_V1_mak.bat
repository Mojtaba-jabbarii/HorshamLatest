call "C:\Program Files (x86)\Intel\oneAPI\compiler\2022.1.0\windows\..\..\..\setvars.bat"  ia32
pushd "C:\GitHub\SF_BESS_Horsham\PSCAD_sim\base_model\20240718_HSFBESS_V1_FW10\HSFBESS_V1.if18_x86\"
nmake -f HSFBESS_V1.mak
popd
