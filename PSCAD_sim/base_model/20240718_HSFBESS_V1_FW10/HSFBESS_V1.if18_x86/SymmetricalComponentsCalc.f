!=======================================================================
! Generated by: PSCAD v5.0.1.0
! Warning:  The content of this file is automatically generated.
!           Do not modify, as any changes made here will be lost!
!-----------------------------------------------------------------------
! Component     : SymmetricalComponentsCalc
! Description   : 
!-----------------------------------------------------------------------


!=======================================================================

      SUBROUTINE SymmetricalComponentsCalcDyn(neg, pos, zero, C, B, A,  &
     &   Frequency)

!---------------------------------------
! Standard includes 
!---------------------------------------

      INCLUDE 'nd.h'
      INCLUDE 'emtconst.h'
      INCLUDE 'emtstor.h'
      INCLUDE 's0.h'
      INCLUDE 's1.h'
      INCLUDE 's2.h'
      INCLUDE 's4.h'
      INCLUDE 'branches.h'
      INCLUDE 'pscadv3.h'
      INCLUDE 'fnames.h'
      INCLUDE 'radiolinks.h'
      INCLUDE 'matlab.h'
      INCLUDE 'rtconfig.h'

!---------------------------------------
! Function/Subroutine Declarations 
!---------------------------------------

!     SUBR    FTN180        ! FFT Calculation

!---------------------------------------
! Variable Declarations 
!---------------------------------------


! Subroutine Arguments
      REAL,    INTENT(IN)  :: C, B, A, Frequency
      REAL,    INTENT(OUT) :: neg, pos, zero

! Electrical Node Indices

! Control Signals
      REAL     RT_1, RT_2(7), RT_3(7), RT_4
      REAL     RT_5(7), RT_6(7), magA, RT_7
      REAL     RT_8(7), RT_9(7), magB, magC
      REAL     RT_10, RT_11, RT_12, RT_13, RT_14
      REAL     RT_15, RT_16, RT_17, RT_18
      REAL     A_real_neg, A_imag_neg, RT_19
      REAL     RT_20, RT_21, C_real_neg
      REAL     C_imag_neg, B_imag_neg, RT_22
      REAL     B_real_neg, RT_23, RT_24, RT_25
      REAL     RT_26, RT_27, RT_28, RT_29, RT_30
      REAL     RT_31, RT_32, RT_33, RT_34, RT_35
      REAL     RT_36, A_real_pos, B_real_pos
      REAL     C_real_pos, RT_37, RT_38, RT_39
      REAL     RT_40, RT_41, RT_42, RT_43
      REAL     A_imag_pos, B_imag_pos, C_imag_pos
      REAL     RT_44, RT_45, RT_46, RT_47, RT_48
      REAL     RT_49, RT_50, RT_51, phaseA, RT_52
      REAL     A_real_zer, A_imag_zer, RT_53
      REAL     phaseC, RT_54, C_real_zer
      REAL     C_imag_zer, B_imag_zer, RT_55
      REAL     B_real_zer, RT_56, phaseB

! Internal Variables
      LOGICAL  LVD1_1
      INTEGER  IVD1_1

! Indexing variables
      INTEGER ICALL_NO                            ! Module call num
      INTEGER ISTOF, IT_0                         ! Storage Indices
      INTEGER IPGB                                ! Control/Monitoring
      INTEGER ISUBS                               ! SS/Node/Branch/Xfmr


!---------------------------------------
! Local Indices 
!---------------------------------------

! Dsdyn <-> Dsout transfer index storage

      NTXFR = NTXFR + 1

      TXFR(NTXFR,1) = NSTOL
      TXFR(NTXFR,2) = NSTOI
      TXFR(NTXFR,3) = NSTOF
      TXFR(NTXFR,4) = NSTOC

! Increment and assign runtime configuration call indices

      ICALL_NO  = NCALL_NO
      NCALL_NO  = NCALL_NO + 1

! Increment global storage indices

      ISTOF     = NSTOF
      NSTOF     = NSTOF + 123
      IPGB      = NPGB
      NPGB      = NPGB + 9
      NNODE     = NNODE + 2
      NCSCS     = NCSCS + 0
      NCSCR     = NCSCR + 0

!---------------------------------------
! Transfers from storage arrays 
!---------------------------------------

      neg      = STOF(ISTOF + 1)
      pos      = STOF(ISTOF + 2)
      zero     = STOF(ISTOF + 3)
      RT_1     = STOF(ISTOF + 8)
      RT_4     = STOF(ISTOF + 23)
      magA     = STOF(ISTOF + 38)
      RT_7     = STOF(ISTOF + 39)
      magB     = STOF(ISTOF + 54)
      magC     = STOF(ISTOF + 55)
      RT_10    = STOF(ISTOF + 56)
      RT_11    = STOF(ISTOF + 57)
      RT_12    = STOF(ISTOF + 58)
      RT_13    = STOF(ISTOF + 59)
      RT_14    = STOF(ISTOF + 60)
      RT_15    = STOF(ISTOF + 61)
      RT_16    = STOF(ISTOF + 62)
      RT_17    = STOF(ISTOF + 63)
      RT_18    = STOF(ISTOF + 64)
      A_real_neg = STOF(ISTOF + 65)
      A_imag_neg = STOF(ISTOF + 66)
      RT_19    = STOF(ISTOF + 67)
      RT_20    = STOF(ISTOF + 68)
      RT_21    = STOF(ISTOF + 69)
      C_real_neg = STOF(ISTOF + 70)
      C_imag_neg = STOF(ISTOF + 71)
      B_imag_neg = STOF(ISTOF + 72)
      RT_22    = STOF(ISTOF + 73)
      B_real_neg = STOF(ISTOF + 74)
      RT_23    = STOF(ISTOF + 75)
      RT_24    = STOF(ISTOF + 76)
      RT_25    = STOF(ISTOF + 77)
      RT_26    = STOF(ISTOF + 78)
      RT_27    = STOF(ISTOF + 79)
      RT_28    = STOF(ISTOF + 80)
      RT_29    = STOF(ISTOF + 81)
      RT_30    = STOF(ISTOF + 82)
      RT_31    = STOF(ISTOF + 83)
      RT_32    = STOF(ISTOF + 84)
      RT_33    = STOF(ISTOF + 85)
      RT_34    = STOF(ISTOF + 86)
      RT_35    = STOF(ISTOF + 87)
      RT_36    = STOF(ISTOF + 88)
      A_real_pos = STOF(ISTOF + 89)
      B_real_pos = STOF(ISTOF + 90)
      C_real_pos = STOF(ISTOF + 91)
      RT_37    = STOF(ISTOF + 92)
      RT_38    = STOF(ISTOF + 93)
      RT_39    = STOF(ISTOF + 94)
      RT_40    = STOF(ISTOF + 95)
      RT_41    = STOF(ISTOF + 96)
      RT_42    = STOF(ISTOF + 97)
      RT_43    = STOF(ISTOF + 98)
      A_imag_pos = STOF(ISTOF + 99)
      B_imag_pos = STOF(ISTOF + 100)
      C_imag_pos = STOF(ISTOF + 101)
      RT_44    = STOF(ISTOF + 102)
      RT_45    = STOF(ISTOF + 103)
      RT_46    = STOF(ISTOF + 104)
      RT_47    = STOF(ISTOF + 105)
      RT_48    = STOF(ISTOF + 106)
      RT_49    = STOF(ISTOF + 107)
      RT_50    = STOF(ISTOF + 108)
      RT_51    = STOF(ISTOF + 109)
      phaseA   = STOF(ISTOF + 110)
      RT_52    = STOF(ISTOF + 111)
      A_real_zer = STOF(ISTOF + 112)
      A_imag_zer = STOF(ISTOF + 113)
      RT_53    = STOF(ISTOF + 114)
      phaseC   = STOF(ISTOF + 115)
      RT_54    = STOF(ISTOF + 116)
      C_real_zer = STOF(ISTOF + 117)
      C_imag_zer = STOF(ISTOF + 118)
      B_imag_zer = STOF(ISTOF + 119)
      RT_55    = STOF(ISTOF + 120)
      B_real_zer = STOF(ISTOF + 121)
      RT_56    = STOF(ISTOF + 122)
      phaseB   = STOF(ISTOF + 123)

! Array (1:7) quantities...
      DO IT_0 = 1,7
         RT_2(IT_0) = STOF(ISTOF + 8 + IT_0)
         RT_3(IT_0) = STOF(ISTOF + 15 + IT_0)
         RT_5(IT_0) = STOF(ISTOF + 23 + IT_0)
         RT_6(IT_0) = STOF(ISTOF + 30 + IT_0)
         RT_8(IT_0) = STOF(ISTOF + 39 + IT_0)
         RT_9(IT_0) = STOF(ISTOF + 46 + IT_0)
      END DO


!---------------------------------------
! Electrical Node Lookup 
!---------------------------------------


!---------------------------------------
! Configuration of Models 
!---------------------------------------

      IF ( TIMEZERO ) THEN
         FILENAME = 'SymmetricalComponentsCalc.dta'
         CALL EMTDC_OPENFILE
         SECTION = 'DATADSD:'
         CALL EMTDC_GOTOSECTION
      ENDIF
!---------------------------------------
! Generated code from module definition 
!---------------------------------------


! 20:[const] Real Constant 
      RT_46 = 3.0

! 40:[fft] On-Line Frequency Scanner 
      IVD1_1=0
      CALL COMPONENT_ID(ICALL_NO,95111471)
      CALL FTN180(0,0,7,1,Frequency,Frequency,A,IVD1_1,RT_8,RT_9,RT_7)
!

! 60:[fft] On-Line Frequency Scanner 
      IVD1_1=0
      CALL COMPONENT_ID(ICALL_NO,1761129337)
      CALL FTN180(0,0,7,1,Frequency,Frequency,B,IVD1_1,RT_5,RT_6,RT_4)
!

! 70:[const] Real Constant 
      RT_29 = 2.0943951

! 80:[const] Real Constant 
      RT_42 = 3.0

! 100:[const] Real Constant 
      RT_27 = 4.1887902

! 110:[fft] On-Line Frequency Scanner 
      IVD1_1=0
      CALL COMPONENT_ID(ICALL_NO,2049775284)
      CALL FTN180(0,0,7,1,Frequency,Frequency,C,IVD1_1,RT_2,RT_3,RT_1)
!

! 120:[const] Real Constant 
      RT_25 = 4.1887902

! 130:[const] Real Constant 
      RT_12 = 3.0

! 140:[const] Real Constant 
      RT_26 = 2.0943951

! 150:[datatap] Scalar/Array Tap 
      magA = RT_8(1)

! 160:[datatap] Scalar/Array Tap 
      phaseA = RT_9(1)

! 170:[trig] Trigonometric Functions 
!  Trig-Func
      RT_36 = COS(phaseA)
!

! 180:[gain] Gain Block 
!  Gain
      A_real_pos = magA * RT_36

! 190:[datatap] Scalar/Array Tap 
      magB = RT_5(1)

! 200:[trig] Trigonometric Functions 
!  Trig-Func
      RT_35 = SIN(phaseA)
!

! 210:[gain] Gain Block 
!  Gain
      A_imag_pos = magA * RT_35

! 220:[datatap] Scalar/Array Tap 
      phaseB = RT_6(1)

! 230:[sumjct] Summing/Differencing Junctions 
      RT_30 = + phaseB + RT_29

! 240:[trig] Trigonometric Functions 
!  Trig-Func
      RT_32 = COS(RT_30)
!

! 250:[gain] Gain Block 
!  Gain
      B_real_pos = magB * RT_32

! 260:[trig] Trigonometric Functions 
!  Trig-Func
      RT_31 = SIN(RT_30)
!

! 270:[gain] Gain Block 
!  Gain
      B_imag_pos = magB * RT_31

! 280:[datatap] Scalar/Array Tap 
      magC = RT_2(1)

! 290:[datatap] Scalar/Array Tap 
      phaseC = RT_3(1)

! 300:[trig] Trigonometric Functions 
!  Trig-Func
      RT_18 = COS(phaseA)
!

! 310:[gain] Gain Block 
!  Gain
      A_real_neg = magA * RT_18

! 320:[trig] Trigonometric Functions 
!  Trig-Func
      RT_17 = SIN(phaseA)
!

! 330:[gain] Gain Block 
!  Gain
      A_imag_neg = magA * RT_17

! 340:[sumjct] Summing/Differencing Junctions 
      RT_24 = + phaseB + RT_25

! 350:[trig] Trigonometric Functions 
!  Trig-Func
      RT_23 = COS(RT_24)
!

! 360:[gain] Gain Block 
!  Gain
      B_real_neg = magB * RT_23

! 370:[trig] Trigonometric Functions 
!  Trig-Func
      RT_22 = SIN(RT_24)
!

! 380:[gain] Gain Block 
!  Gain
      B_imag_neg = magB * RT_22

! 390:[sumjct] Summing/Differencing Junctions 
      RT_20 = + phaseC + RT_26

! 400:[trig] Trigonometric Functions 
!  Trig-Func
      RT_21 = COS(RT_20)
!

! 410:[gain] Gain Block 
!  Gain
      C_real_neg = magC * RT_21

! 420:[trig] Trigonometric Functions 
!  Trig-Func
      RT_19 = SIN(RT_20)
!

! 430:[gain] Gain Block 
!  Gain
      C_imag_neg = magC * RT_19

! 440:[trig] Trigonometric Functions 
!  Trig-Func
      RT_52 = COS(phaseA)
!

! 450:[gain] Gain Block 
!  Gain
      A_real_zer = magA * RT_52

! 460:[trig] Trigonometric Functions 
!  Trig-Func
      RT_51 = SIN(phaseA)
!

! 470:[gain] Gain Block 
!  Gain
      A_imag_zer = magA * RT_51

! 480:[trig] Trigonometric Functions 
!  Trig-Func
      RT_56 = COS(phaseB)
!

! 490:[gain] Gain Block 
!  Gain
      B_real_zer = magB * RT_56

! 500:[trig] Trigonometric Functions 
!  Trig-Func
      RT_55 = SIN(phaseB)
!

! 510:[gain] Gain Block 
!  Gain
      B_imag_zer = magB * RT_55

! 520:[trig] Trigonometric Functions 
!  Trig-Func
      RT_54 = COS(phaseC)
!

! 530:[gain] Gain Block 
!  Gain
      C_real_zer = magC * RT_54

! 540:[trig] Trigonometric Functions 
!  Trig-Func
      RT_53 = SIN(phaseC)
!

! 550:[gain] Gain Block 
!  Gain
      C_imag_zer = magC * RT_53

! 560:[sumjct] Summing/Differencing Junctions 
      RT_28 = + phaseC + RT_27

! 570:[trig] Trigonometric Functions 
!  Trig-Func
      RT_34 = COS(RT_28)
!

! 580:[gain] Gain Block 
!  Gain
      C_real_pos = magC * RT_34

! 590:[trig] Trigonometric Functions 
!  Trig-Func
      RT_33 = SIN(RT_28)
!

! 600:[gain] Gain Block 
!  Gain
      C_imag_pos = magC * RT_33

! 610:[sumjct] Summing/Differencing Junctions 
      RT_16 = + A_real_neg + B_real_neg + C_real_neg

! 620:[square] Square 
      RT_15 = RT_16 * RT_16

! 630:[sumjct] Summing/Differencing Junctions 
      RT_10 = + A_imag_neg + B_imag_neg + C_imag_neg

! 640:[square] Square 
      RT_11 = RT_10 * RT_10

! 650:[sumjct] Summing/Differencing Junctions 
      RT_50 = + A_real_zer + B_real_zer + C_real_zer

! 660:[square] Square 
      RT_49 = RT_50 * RT_50

! 670:[sumjct] Summing/Differencing Junctions 
      RT_44 = + A_imag_zer + B_imag_zer + C_imag_zer

! 680:[square] Square 
      RT_45 = RT_44 * RT_44

! 690:[sumjct] Summing/Differencing Junctions 
      RT_37 = + A_real_pos + B_real_pos + C_real_pos

! 700:[square] Square 
      RT_38 = RT_37 * RT_37

! 710:[sumjct] Summing/Differencing Junctions 
      RT_43 = + A_imag_pos + B_imag_pos + C_imag_pos

! 720:[square] Square 
      RT_39 = RT_43 * RT_43

! 730:[sumjct] Summing/Differencing Junctions 
      RT_14 = + RT_15 + RT_11

! 740:[sumjct] Summing/Differencing Junctions 
      RT_40 = + RT_38 + RT_39

! 750:[sumjct] Summing/Differencing Junctions 
      RT_48 = + RT_49 + RT_45

! 760:[sqrt] Square Root 
      LVD1_1 = STORL(NSTORL)
      IF (RT_14 .LT. 0.0) THEN
        RT_13 = 0.0
        IF (.NOT. LVD1_1) THEN
          CALL EMTDC_MESSAGE(ICALL_NO,1910578618,1,2,"A negative value i&
     &s detected as an input to the Square Root function. ")
          CALL EMTDC_MESSAGE(ICALL_NO,1910578618,1,-1,"Input is treated &
     &as 0.0.")
          STORL(NSTORL) = .TRUE.
        ENDIF
      ELSE
         RT_13 = SQRT(RT_14)
      ENDIF
      NSTORL = NSTORL + 1

! 770:[sqrt] Square Root 
      LVD1_1 = STORL(NSTORL)
      IF (RT_40 .LT. 0.0) THEN
        RT_41 = 0.0
        IF (.NOT. LVD1_1) THEN
          CALL EMTDC_MESSAGE(ICALL_NO,1357403854,1,2,"A negative value i&
     &s detected as an input to the Square Root function. ")
          CALL EMTDC_MESSAGE(ICALL_NO,1357403854,1,-1,"Input is treated &
     &as 0.0.")
          STORL(NSTORL) = .TRUE.
        ENDIF
      ELSE
         RT_41 = SQRT(RT_40)
      ENDIF
      NSTORL = NSTORL + 1

! 780:[sqrt] Square Root 
      LVD1_1 = STORL(NSTORL)
      IF (RT_48 .LT. 0.0) THEN
        RT_47 = 0.0
        IF (.NOT. LVD1_1) THEN
          CALL EMTDC_MESSAGE(ICALL_NO,1148924766,1,2,"A negative value i&
     &s detected as an input to the Square Root function. ")
          CALL EMTDC_MESSAGE(ICALL_NO,1148924766,1,-1,"Input is treated &
     &as 0.0.")
          STORL(NSTORL) = .TRUE.
        ENDIF
      ELSE
         RT_47 = SQRT(RT_48)
      ENDIF
      NSTORL = NSTORL + 1

! 790:[div] Divider 
      IF (ABS(RT_12) .LT. 1.0E-100) THEN
         IF (RT_12 .LT. 0.0)  THEN
            neg = -1.0E100 * RT_13
         ELSE
            neg =  1.0E100 * RT_13
         ENDIF
      ELSE
         neg = RT_13 / RT_12
      ENDIF

! 800:[div] Divider 
      IF (ABS(RT_42) .LT. 1.0E-100) THEN
         IF (RT_42 .LT. 0.0)  THEN
            pos = -1.0E100 * RT_41
         ELSE
            pos =  1.0E100 * RT_41
         ENDIF
      ELSE
         pos = RT_41 / RT_42
      ENDIF

! 810:[div] Divider 
      IF (ABS(RT_46) .LT. 1.0E-100) THEN
         IF (RT_46 .LT. 0.0)  THEN
            zero = -1.0E100 * RT_47
         ELSE
            zero =  1.0E100 * RT_47
         ENDIF
      ELSE
         zero = RT_47 / RT_46
      ENDIF

! 830:[pgb] Output Channel 'neg'

      PGB(IPGB+1) = neg

! 850:[pgb] Output Channel 'phaseC'

      PGB(IPGB+2) = phaseC

! 860:[pgb] Output Channel 'pos'

      PGB(IPGB+3) = pos

! 870:[pgb] Output Channel 'magC'

      PGB(IPGB+4) = magC

! 880:[pgb] Output Channel 'phaseB'

      PGB(IPGB+5) = phaseB

! 890:[pgb] Output Channel 'magB'

      PGB(IPGB+6) = magB

! 900:[pgb] Output Channel 'phaseA'

      PGB(IPGB+7) = phaseA

! 920:[pgb] Output Channel 'magA'

      PGB(IPGB+8) = magA

! 930:[pgb] Output Channel 'zero'

      PGB(IPGB+9) = zero

!---------------------------------------
! Feedbacks and transfers to storage 
!---------------------------------------

      STOF(ISTOF + 1) = neg
      STOF(ISTOF + 2) = pos
      STOF(ISTOF + 3) = zero
      STOF(ISTOF + 4) = C
      STOF(ISTOF + 5) = B
      STOF(ISTOF + 6) = A
      STOF(ISTOF + 7) = Frequency
      STOF(ISTOF + 8) = RT_1
      STOF(ISTOF + 23) = RT_4
      STOF(ISTOF + 38) = magA
      STOF(ISTOF + 39) = RT_7
      STOF(ISTOF + 54) = magB
      STOF(ISTOF + 55) = magC
      STOF(ISTOF + 56) = RT_10
      STOF(ISTOF + 57) = RT_11
      STOF(ISTOF + 58) = RT_12
      STOF(ISTOF + 59) = RT_13
      STOF(ISTOF + 60) = RT_14
      STOF(ISTOF + 61) = RT_15
      STOF(ISTOF + 62) = RT_16
      STOF(ISTOF + 63) = RT_17
      STOF(ISTOF + 64) = RT_18
      STOF(ISTOF + 65) = A_real_neg
      STOF(ISTOF + 66) = A_imag_neg
      STOF(ISTOF + 67) = RT_19
      STOF(ISTOF + 68) = RT_20
      STOF(ISTOF + 69) = RT_21
      STOF(ISTOF + 70) = C_real_neg
      STOF(ISTOF + 71) = C_imag_neg
      STOF(ISTOF + 72) = B_imag_neg
      STOF(ISTOF + 73) = RT_22
      STOF(ISTOF + 74) = B_real_neg
      STOF(ISTOF + 75) = RT_23
      STOF(ISTOF + 76) = RT_24
      STOF(ISTOF + 77) = RT_25
      STOF(ISTOF + 78) = RT_26
      STOF(ISTOF + 79) = RT_27
      STOF(ISTOF + 80) = RT_28
      STOF(ISTOF + 81) = RT_29
      STOF(ISTOF + 82) = RT_30
      STOF(ISTOF + 83) = RT_31
      STOF(ISTOF + 84) = RT_32
      STOF(ISTOF + 85) = RT_33
      STOF(ISTOF + 86) = RT_34
      STOF(ISTOF + 87) = RT_35
      STOF(ISTOF + 88) = RT_36
      STOF(ISTOF + 89) = A_real_pos
      STOF(ISTOF + 90) = B_real_pos
      STOF(ISTOF + 91) = C_real_pos
      STOF(ISTOF + 92) = RT_37
      STOF(ISTOF + 93) = RT_38
      STOF(ISTOF + 94) = RT_39
      STOF(ISTOF + 95) = RT_40
      STOF(ISTOF + 96) = RT_41
      STOF(ISTOF + 97) = RT_42
      STOF(ISTOF + 98) = RT_43
      STOF(ISTOF + 99) = A_imag_pos
      STOF(ISTOF + 100) = B_imag_pos
      STOF(ISTOF + 101) = C_imag_pos
      STOF(ISTOF + 102) = RT_44
      STOF(ISTOF + 103) = RT_45
      STOF(ISTOF + 104) = RT_46
      STOF(ISTOF + 105) = RT_47
      STOF(ISTOF + 106) = RT_48
      STOF(ISTOF + 107) = RT_49
      STOF(ISTOF + 108) = RT_50
      STOF(ISTOF + 109) = RT_51
      STOF(ISTOF + 110) = phaseA
      STOF(ISTOF + 111) = RT_52
      STOF(ISTOF + 112) = A_real_zer
      STOF(ISTOF + 113) = A_imag_zer
      STOF(ISTOF + 114) = RT_53
      STOF(ISTOF + 115) = phaseC
      STOF(ISTOF + 116) = RT_54
      STOF(ISTOF + 117) = C_real_zer
      STOF(ISTOF + 118) = C_imag_zer
      STOF(ISTOF + 119) = B_imag_zer
      STOF(ISTOF + 120) = RT_55
      STOF(ISTOF + 121) = B_real_zer
      STOF(ISTOF + 122) = RT_56
      STOF(ISTOF + 123) = phaseB

! Array (1:7) quantities...
      DO IT_0 = 1,7
         STOF(ISTOF + 8 + IT_0) = RT_2(IT_0)
         STOF(ISTOF + 15 + IT_0) = RT_3(IT_0)
         STOF(ISTOF + 23 + IT_0) = RT_5(IT_0)
         STOF(ISTOF + 30 + IT_0) = RT_6(IT_0)
         STOF(ISTOF + 39 + IT_0) = RT_8(IT_0)
         STOF(ISTOF + 46 + IT_0) = RT_9(IT_0)
      END DO


!---------------------------------------
! Transfer to Exports
!---------------------------------------
      !neg      is output
      !pos      is output
      !zero     is output

!---------------------------------------
! Close Model Data read 
!---------------------------------------

      IF ( TIMEZERO ) CALL EMTDC_CLOSEFILE
      RETURN
      END

!=======================================================================

      SUBROUTINE SymmetricalComponentsCalcOut()

!---------------------------------------
! Standard includes 
!---------------------------------------

      INCLUDE 'nd.h'
      INCLUDE 'emtconst.h'
      INCLUDE 'emtstor.h'
      INCLUDE 's0.h'
      INCLUDE 's1.h'
      INCLUDE 's2.h'
      INCLUDE 's4.h'
      INCLUDE 'branches.h'
      INCLUDE 'pscadv3.h'
      INCLUDE 'fnames.h'
      INCLUDE 'radiolinks.h'
      INCLUDE 'matlab.h'
      INCLUDE 'rtconfig.h'

!---------------------------------------
! Function/Subroutine Declarations 
!---------------------------------------


!---------------------------------------
! Variable Declarations 
!---------------------------------------


! Electrical Node Indices

! Control Signals
      REAL     RT_12, RT_25, RT_26, RT_27, RT_29
      REAL     RT_42, RT_46

! Internal Variables

! Indexing variables
      INTEGER ICALL_NO                            ! Module call num
      INTEGER ISTOL, ISTOI, ISTOF, ISTOC          ! Storage Indices
      INTEGER ISUBS                               ! SS/Node/Branch/Xfmr


!---------------------------------------
! Local Indices 
!---------------------------------------

! Dsdyn <-> Dsout transfer index storage

      NTXFR = NTXFR + 1

      ISTOL = TXFR(NTXFR,1)
      ISTOI = TXFR(NTXFR,2)
      ISTOF = TXFR(NTXFR,3)
      ISTOC = TXFR(NTXFR,4)

! Increment and assign runtime configuration call indices

      ICALL_NO  = NCALL_NO
      NCALL_NO  = NCALL_NO + 1

! Increment global storage indices

      NPGB      = NPGB + 9
      NNODE     = NNODE + 2
      NCSCS     = NCSCS + 0
      NCSCR     = NCSCR + 0

!---------------------------------------
! Transfers from storage arrays 
!---------------------------------------

      RT_12    = STOF(ISTOF + 58)
      RT_25    = STOF(ISTOF + 77)
      RT_26    = STOF(ISTOF + 78)
      RT_27    = STOF(ISTOF + 79)
      RT_29    = STOF(ISTOF + 81)
      RT_42    = STOF(ISTOF + 97)
      RT_46    = STOF(ISTOF + 104)


!---------------------------------------
! Electrical Node Lookup 
!---------------------------------------


!---------------------------------------
! Configuration of Models 
!---------------------------------------

      IF ( TIMEZERO ) THEN
         FILENAME = 'SymmetricalComponentsCalc.dta'
         CALL EMTDC_OPENFILE
         SECTION = 'DATADSO:'
         CALL EMTDC_GOTOSECTION
      ENDIF
!---------------------------------------
! Generated code from module definition 
!---------------------------------------


! 20:[const] Real Constant 

      RT_46 = 3.0

! 70:[const] Real Constant 

      RT_29 = 2.0943951

! 80:[const] Real Constant 

      RT_42 = 3.0

! 100:[const] Real Constant 

      RT_27 = 4.1887902

! 120:[const] Real Constant 

      RT_25 = 4.1887902

! 130:[const] Real Constant 

      RT_12 = 3.0

! 140:[const] Real Constant 

      RT_26 = 2.0943951

!---------------------------------------
! Feedbacks and transfers to storage 
!---------------------------------------

      STOF(ISTOF + 58) = RT_12
      STOF(ISTOF + 77) = RT_25
      STOF(ISTOF + 78) = RT_26
      STOF(ISTOF + 79) = RT_27
      STOF(ISTOF + 81) = RT_29
      STOF(ISTOF + 97) = RT_42
      STOF(ISTOF + 104) = RT_46


!---------------------------------------
! Close Model Data read 
!---------------------------------------

      IF ( TIMEZERO ) CALL EMTDC_CLOSEFILE
      RETURN
      END

!=======================================================================

      SUBROUTINE SymmetricalComponentsCalcDyn_Begin(Frequency)

!---------------------------------------
! Standard includes 
!---------------------------------------

      INCLUDE 'nd.h'
      INCLUDE 'emtconst.h'
      INCLUDE 's0.h'
      INCLUDE 's1.h'
      INCLUDE 's4.h'
      INCLUDE 'branches.h'
      INCLUDE 'pscadv3.h'
      INCLUDE 'radiolinks.h'
      INCLUDE 'rtconfig.h'

!---------------------------------------
! Function/Subroutine Declarations 
!---------------------------------------


!---------------------------------------
! Variable Declarations 
!---------------------------------------


! Subroutine Arguments
      REAL,    INTENT(IN)  :: Frequency

! Electrical Node Indices

! Control Signals
      REAL     RT_12, RT_25, RT_26, RT_27, RT_29
      REAL     RT_42, RT_46

! Internal Variables

! Indexing variables
      INTEGER ICALL_NO                            ! Module call num
      INTEGER ISUBS                               ! SS/Node/Branch/Xfmr


!---------------------------------------
! Local Indices 
!---------------------------------------


! Increment and assign runtime configuration call indices

      ICALL_NO  = NCALL_NO
      NCALL_NO  = NCALL_NO + 1

! Increment global storage indices

      NNODE     = NNODE + 2
      NCSCS     = NCSCS + 0
      NCSCR     = NCSCR + 0

!---------------------------------------
! Electrical Node Lookup 
!---------------------------------------


!---------------------------------------
! Generated code from module definition 
!---------------------------------------


! 20:[const] Real Constant 
      RT_46 = 3.0

! 40:[fft] On-Line Frequency Scanner 

! 60:[fft] On-Line Frequency Scanner 

! 70:[const] Real Constant 
      RT_29 = 2.0943951

! 80:[const] Real Constant 
      RT_42 = 3.0

! 100:[const] Real Constant 
      RT_27 = 4.1887902

! 110:[fft] On-Line Frequency Scanner 

! 120:[const] Real Constant 
      RT_25 = 4.1887902

! 130:[const] Real Constant 
      RT_12 = 3.0

! 140:[const] Real Constant 
      RT_26 = 2.0943951

! 150:[datatap] Scalar/Array Tap 

! 160:[datatap] Scalar/Array Tap 

! 170:[trig] Trigonometric Functions 

! 180:[gain] Gain Block 

! 190:[datatap] Scalar/Array Tap 

! 200:[trig] Trigonometric Functions 

! 210:[gain] Gain Block 

! 220:[datatap] Scalar/Array Tap 

! 230:[sumjct] Summing/Differencing Junctions 

! 240:[trig] Trigonometric Functions 

! 250:[gain] Gain Block 

! 260:[trig] Trigonometric Functions 

! 270:[gain] Gain Block 

! 280:[datatap] Scalar/Array Tap 

! 290:[datatap] Scalar/Array Tap 

! 300:[trig] Trigonometric Functions 

! 310:[gain] Gain Block 

! 320:[trig] Trigonometric Functions 

! 330:[gain] Gain Block 

! 340:[sumjct] Summing/Differencing Junctions 

! 350:[trig] Trigonometric Functions 

! 360:[gain] Gain Block 

! 370:[trig] Trigonometric Functions 

! 380:[gain] Gain Block 

! 390:[sumjct] Summing/Differencing Junctions 

! 400:[trig] Trigonometric Functions 

! 410:[gain] Gain Block 

! 420:[trig] Trigonometric Functions 

! 430:[gain] Gain Block 

! 440:[trig] Trigonometric Functions 

! 450:[gain] Gain Block 

! 460:[trig] Trigonometric Functions 

! 470:[gain] Gain Block 

! 480:[trig] Trigonometric Functions 

! 490:[gain] Gain Block 

! 500:[trig] Trigonometric Functions 

! 510:[gain] Gain Block 

! 520:[trig] Trigonometric Functions 

! 530:[gain] Gain Block 

! 540:[trig] Trigonometric Functions 

! 550:[gain] Gain Block 

! 560:[sumjct] Summing/Differencing Junctions 

! 570:[trig] Trigonometric Functions 

! 580:[gain] Gain Block 

! 590:[trig] Trigonometric Functions 

! 600:[gain] Gain Block 

! 610:[sumjct] Summing/Differencing Junctions 

! 620:[square] Square 

! 630:[sumjct] Summing/Differencing Junctions 

! 640:[square] Square 

! 650:[sumjct] Summing/Differencing Junctions 

! 660:[square] Square 

! 670:[sumjct] Summing/Differencing Junctions 

! 680:[square] Square 

! 690:[sumjct] Summing/Differencing Junctions 

! 700:[square] Square 

! 710:[sumjct] Summing/Differencing Junctions 

! 720:[square] Square 

! 730:[sumjct] Summing/Differencing Junctions 

! 740:[sumjct] Summing/Differencing Junctions 

! 750:[sumjct] Summing/Differencing Junctions 

! 760:[sqrt] Square Root 

! 770:[sqrt] Square Root 

! 780:[sqrt] Square Root 

! 790:[div] Divider 

! 800:[div] Divider 

! 810:[div] Divider 

! 830:[pgb] Output Channel 'neg'

! 850:[pgb] Output Channel 'phaseC'

! 860:[pgb] Output Channel 'pos'

! 870:[pgb] Output Channel 'magC'

! 880:[pgb] Output Channel 'phaseB'

! 890:[pgb] Output Channel 'magB'

! 900:[pgb] Output Channel 'phaseA'

! 920:[pgb] Output Channel 'magA'

! 930:[pgb] Output Channel 'zero'

      RETURN
      END

!=======================================================================

      SUBROUTINE SymmetricalComponentsCalcOut_Begin(Frequency)

!---------------------------------------
! Standard includes 
!---------------------------------------

      INCLUDE 'nd.h'
      INCLUDE 'emtconst.h'
      INCLUDE 's0.h'
      INCLUDE 's1.h'
      INCLUDE 's4.h'
      INCLUDE 'branches.h'
      INCLUDE 'pscadv3.h'
      INCLUDE 'radiolinks.h'
      INCLUDE 'rtconfig.h'

!---------------------------------------
! Function/Subroutine Declarations 
!---------------------------------------


!---------------------------------------
! Variable Declarations 
!---------------------------------------


! Subroutine Arguments
      REAL,    INTENT(IN)  :: Frequency

! Electrical Node Indices

! Control Signals
      REAL     RT_12, RT_25, RT_26, RT_27, RT_29
      REAL     RT_42, RT_46

! Internal Variables

! Indexing variables
      INTEGER ICALL_NO                            ! Module call num
      INTEGER ISUBS                               ! SS/Node/Branch/Xfmr


!---------------------------------------
! Local Indices 
!---------------------------------------


! Increment and assign runtime configuration call indices

      ICALL_NO  = NCALL_NO
      NCALL_NO  = NCALL_NO + 1

! Increment global storage indices

      NNODE     = NNODE + 2
      NCSCS     = NCSCS + 0
      NCSCR     = NCSCR + 0

!---------------------------------------
! Electrical Node Lookup 
!---------------------------------------


!---------------------------------------
! Generated code from module definition 
!---------------------------------------


! 20:[const] Real Constant 
      RT_46 = 3.0

! 70:[const] Real Constant 
      RT_29 = 2.0943951

! 80:[const] Real Constant 
      RT_42 = 3.0

! 100:[const] Real Constant 
      RT_27 = 4.1887902

! 120:[const] Real Constant 
      RT_25 = 4.1887902

! 130:[const] Real Constant 
      RT_12 = 3.0

! 140:[const] Real Constant 
      RT_26 = 2.0943951

      RETURN
      END
