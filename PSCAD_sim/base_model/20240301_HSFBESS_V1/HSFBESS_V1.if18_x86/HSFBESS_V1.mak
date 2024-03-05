
#------------------------------------------------------------------------------
# Project 'HSFBESS_V1' make using the 'Intel_ Fortran Compiler Classic 2021.6.0' compiler.
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# All project
#------------------------------------------------------------------------------

all: targets
	@echo !--Make: succeeded.



#------------------------------------------------------------------------------
# Directories, Platform, and Version
#------------------------------------------------------------------------------

Arch        = windows
EmtdcDir    = C:\Program Files (x86)\PSCAD50\emtdc\if18_x86
EmtdcInc    = $(EmtdcDir)\inc
EmtdcBin    = $(EmtdcDir)\$(Arch)
EmtdcMain   = $(EmtdcBin)\main.obj
EmtdcLib    = $(EmtdcBin)\emtdc.lib
SolverLib    = $(EmtdcBin)\Solver.lib


#------------------------------------------------------------------------------
# Fortran Compiler
#------------------------------------------------------------------------------

FC_Name         = ifort.exe
FC_Suffix       = obj
FC_Args         = /nologo /c /free /real_size:64 /fpconstant /warn:declarations /iface:default /align:dcommons /fpe:0
FC_Debug        =  /O2
FC_Preprocess   = 
FC_Preproswitch = 
FC_Warn         = 
FC_Checks       = 
FC_Includes     = /include:"$(EmtdcInc)" /include:"$(EmtdcDir)" /include:"$(EmtdcBin)"
FC_Compile      = $(FC_Name) $(FC_Args) $(FC_Includes) $(FC_Debug) $(FC_Warn) $(FC_Checks)

#------------------------------------------------------------------------------
# C Compiler
#------------------------------------------------------------------------------

CC_Name     = cl.exe
CC_Suffix   = obj
CC_Args     = /nologo /MT /W3 /EHsc /c
CC_Debug    =  /O2
CC_Includes = 
CC_Compile  = $(CC_Name) $(CC_Args) $(CC_Includes) $(CC_Debug)

#------------------------------------------------------------------------------
# Linker
#------------------------------------------------------------------------------

Link_Name   = link.exe
Link_Debug  = 
Link_Args   = /out:$@ /nologo /nodefaultlib:libc.lib /nodefaultlib:libcmtd.lib /subsystem:console
Link        = $(Link_Name) $(Link_Args) $(Link_Debug)

#------------------------------------------------------------------------------
# Build rules for generated files
#------------------------------------------------------------------------------


.f.$(FC_Suffix):
	@echo !--Compile: $<
	$(FC_Compile) $<



.c.$(CC_Suffix):
	@echo !--Compile: $<
	$(CC_Compile) $<



#------------------------------------------------------------------------------
# Build rules for file references
#------------------------------------------------------------------------------


SMASC_K_090205R03_if18_x86_1.lib: 
	@echo !--Copy: "C:\GitHub\SF_BESS_Horsham\PSCAD_sim\base_model\20240301_HSFBESS_V1\SMASC_K_090205R03\Libs_if18_x86\SMASC_K_090205R03_if18_x86.lib"
	copy "C:\GitHub\SF_BESS_Horsham\PSCAD_sim\base_model\20240301_HSFBESS_V1\SMASC_K_090205R03\Libs_if18_x86\SMASC_K_090205R03_if18_x86.lib" "SMASC_K_090205R03_if18_x86_1.lib"

SMAHYC_021906R01_if18_x86_2.lib: 
	@echo !--Copy: "C:\GitHub\SF_BESS_Horsham\PSCAD_sim\base_model\20240301_HSFBESS_V1\SMAHYC_021906R01\Libs_if18_x86\SMAHYC_021906R01_if18_x86.lib"
	copy "C:\GitHub\SF_BESS_Horsham\PSCAD_sim\base_model\20240301_HSFBESS_V1\SMAHYC_021906R01\Libs_if18_x86\SMAHYC_021906R01_if18_x86.lib" "SMAHYC_021906R01_if18_x86_2.lib"

#------------------------------------------------------------------------------
# Dependencies
#------------------------------------------------------------------------------


FC_Objects = \
 Station.$(FC_Suffix) \
 Main.$(FC_Suffix) \
 InvFbDummy.$(FC_Suffix) \
 DPS_Dummy.$(FC_Suffix) \
 GridSource.$(FC_Suffix) \
 XY_profile.$(FC_Suffix) \
 SignalTruncator.$(FC_Suffix) \
 FaultBlock.$(FC_Suffix) \
 SymmetricalComponentsCalc_1_1.$(FC_Suffix) \
 SymmetricalComponentsCalc.$(FC_Suffix) \
 setpointProfiles.$(FC_Suffix) \
 XY_profile_2.$(FC_Suffix) \
 SignalTruncator_2_3.$(FC_Suffix) \
 TapCtrl_1.$(FC_Suffix) \
 Aggr_Fb_Scaled.$(FC_Suffix) \
 DDSRF_PLL_1.$(FC_Suffix) \
 DDSRF_PLL_2.$(FC_Suffix) \
 HyCtl.$(FC_Suffix) \
 PQ_Select.$(FC_Suffix) \
 SCxxxx.$(FC_Suffix) \
 AvgBridge.$(FC_Suffix) \
 DEBUG_HyCon_Scope.$(FC_Suffix) \
 SC_Scope.$(FC_Suffix)

FC_ObjectsLong = \
 "Station.$(FC_Suffix)" \
 "Main.$(FC_Suffix)" \
 "InvFbDummy.$(FC_Suffix)" \
 "DPS_Dummy.$(FC_Suffix)" \
 "GridSource.$(FC_Suffix)" \
 "XY_profile.$(FC_Suffix)" \
 "SignalTruncator.$(FC_Suffix)" \
 "FaultBlock.$(FC_Suffix)" \
 "SymmetricalComponentsCalc_1_1.$(FC_Suffix)" \
 "SymmetricalComponentsCalc.$(FC_Suffix)" \
 "setpointProfiles.$(FC_Suffix)" \
 "XY_profile_2.$(FC_Suffix)" \
 "SignalTruncator_2_3.$(FC_Suffix)" \
 "TapCtrl_1.$(FC_Suffix)" \
 "Aggr_Fb_Scaled.$(FC_Suffix)" \
 "DDSRF_PLL_1.$(FC_Suffix)" \
 "DDSRF_PLL_2.$(FC_Suffix)" \
 "HyCtl.$(FC_Suffix)" \
 "PQ_Select.$(FC_Suffix)" \
 "SCxxxx.$(FC_Suffix)" \
 "AvgBridge.$(FC_Suffix)" \
 "DEBUG_HyCon_Scope.$(FC_Suffix)" \
 "SC_Scope.$(FC_Suffix)"

CC_Objects =

CC_ObjectsLong =

LK_Objects = \
  SMASC_K_090205R03_if18_x86_1.lib \
  SMAHYC_021906R01_if18_x86_2.lib

LK_ObjectsLong = \
  "SMASC_K_090205R03_if18_x86_1.lib" \
  "SMAHYC_021906R01_if18_x86_2.lib"

SysLibs  = ws2_32.lib

Binary   = HSFBESS_V1.exe

$(Binary): $(FC_Objects) $(CC_Objects) $(LK_Objects) 
	@echo !--Link: $@
	$(Link) "$(EmtdcMain)" $(FC_ObjectsLong) $(CC_ObjectsLong) $(LK_ObjectsLong) "$(EmtdcLib)" "$(SolverLib)" $(SysLibs)

targets: $(Binary)


clean:
	-del EMTDC_V*
	-del *.obj
	-del *.o
	-del *.exe
	@echo !--Make clean: succeeded.



