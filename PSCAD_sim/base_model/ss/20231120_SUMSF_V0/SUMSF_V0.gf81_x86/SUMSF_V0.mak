
#------------------------------------------------------------------------------
# Project 'SUMSF_V0' make using the 'GFortran 8.1' compiler.
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
EmtdcDir    = C:\Program Files (x86)\PSCAD50\emtdc\gf81_x86
EmtdcInc    = $(EmtdcDir)\inc
EmtdcBin    = $(EmtdcDir)\$(Arch)
EmtdcMain   = $(EmtdcBin)\main.obj
EmtdcLib    = $(EmtdcBin)\emtdc.lib
SolverLib    = $(EmtdcBin)\Solver.lib


#------------------------------------------------------------------------------
# Fortran Compiler
#------------------------------------------------------------------------------

FC_Name         = gfortran.exe
FC_Suffix       = o
FC_Args         = -c -ffree-form -fdefault-real-8 -fdefault-double-8
FC_Debug        = -O2
FC_Preprocess   = 
FC_Preproswitch = 
FC_Warn         = -Wconversion
FC_Checks       = 
FC_Includes     = -I"$(EmtdcInc)" -I"$(EmtdcBin)"
FC_Compile      = $(FC_Name) $(FC_Args) $(FC_Includes) $(FC_Debug) $(FC_Warn) $(FC_Checks)

#------------------------------------------------------------------------------
# C Compiler
#------------------------------------------------------------------------------

CC_Name     = gcc.exe
CC_Suffix   = o
CC_Args     = -c
CC_Debug    = -O2
CC_Includes = -I"$(EmtdcInc)" -I"$(EmtdcBin)"
CC_Compile  = $(CC_Name) $(CC_Args) $(CC_Includes) $(CC_Debug)

#------------------------------------------------------------------------------
# Linker
#------------------------------------------------------------------------------

Link_Name   = gcc.exe
Link_Debug  = -O2
Link_Args   = -o $@
Link        = $(Link_Name) $(Link_Args) $(Link_Debug)

#------------------------------------------------------------------------------
# Build rules for generated files
#------------------------------------------------------------------------------


%.$(FC_Suffix): %.f
	@echo !--Compile: $<
	$(FC_Compile) $<


%.$(CC_Suffix): %.c
	@echo !--Compile: $<
	$(CC_Compile) $<



#------------------------------------------------------------------------------
# Build rules for file references
#------------------------------------------------------------------------------


SMASC_K_090205R03_gf81_x86_1.lib: 
	@echo !--Copy: "C:\Users\341510davu\OneDrive - OX2\3. Grid - SUMSF\1. Power System Studies\1. Main Test Environment\Git_SF_BESS_Summerville\PSCAD_sim\base_model\20231120_SUMSF_V0\SMASC_K_090205R03\Libs_gf81_x86\SMASC_K_090205R03_gf81_x86.lib"
	copy "C:\Users\341510davu\OneDrive - OX2\3. Grid - SUMSF\1. Power System Studies\1. Main Test Environment\Git_SF_BESS_Summerville\PSCAD_sim\base_model\20231120_SUMSF_V0\SMASC_K_090205R03\Libs_gf81_x86\SMASC_K_090205R03_gf81_x86.lib" "SMASC_K_090205R03_gf81_x86_1.lib"

SMAHYC_021906R01_gf81_x86_2.lib: 
	@echo !--Copy: "C:\Users\341510davu\OneDrive - OX2\3. Grid - SUMSF\1. Power System Studies\1. Main Test Environment\Git_SF_BESS_Summerville\PSCAD_sim\base_model\20231120_SUMSF_V0\SMAHYC_021906R01\Libs_gf81_x86\SMAHYC_021906R01_gf81_x86.lib"
	copy "C:\Users\341510davu\OneDrive - OX2\3. Grid - SUMSF\1. Power System Studies\1. Main Test Environment\Git_SF_BESS_Summerville\PSCAD_sim\base_model\20231120_SUMSF_V0\SMAHYC_021906R01\Libs_gf81_x86\SMAHYC_021906R01_gf81_x86.lib" "SMAHYC_021906R01_gf81_x86_2.lib"

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
 SCxxxx.$(FC_Suffix) \
 AvgBridge.$(FC_Suffix) \
 Aggr_Fb_Scaled.$(FC_Suffix) \
 DDSRF_PLL_1.$(FC_Suffix) \
 DDSRF_PLL_2.$(FC_Suffix) \
 HyCtl.$(FC_Suffix) \
 PQ_Select.$(FC_Suffix) \
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
 "SCxxxx.$(FC_Suffix)" \
 "AvgBridge.$(FC_Suffix)" \
 "Aggr_Fb_Scaled.$(FC_Suffix)" \
 "DDSRF_PLL_1.$(FC_Suffix)" \
 "DDSRF_PLL_2.$(FC_Suffix)" \
 "HyCtl.$(FC_Suffix)" \
 "PQ_Select.$(FC_Suffix)" \
 "DEBUG_HyCon_Scope.$(FC_Suffix)" \
 "SC_Scope.$(FC_Suffix)"

CC_Objects =

CC_ObjectsLong =

LK_Objects = \
  SMASC_K_090205R03_gf81_x86_1.lib \
  SMAHYC_021906R01_gf81_x86_2.lib

LK_ObjectsLong = \
  "SMASC_K_090205R03_gf81_x86_1.lib" \
  "SMAHYC_021906R01_gf81_x86_2.lib"

SysLibs  = -lgfortran -lstdc++ -lquadmath -lwsock32

Binary   = SUMSF_V0.exe

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



