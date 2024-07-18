!~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
! DLL-Course Demo
!------------------------------------------------------------------------------
! This is a demo that is part of the DLL-Course for creating EMTDC components
! that are primarily executed through a DLL.
!
! Created By:
! ~~~~~~~~~~~
!    PSCAD Design Team <pscad@hvdc.ca>
!    Manitoba HVDC Research Centre Inc.
!    Winnipeg, Manitoba. CANADA
!
!~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    !-----------------------------------------------------------------------------------------------
    ! This code is for linking to PSCAD for an example of how to create wrapper function that are
    ! exported. A DLL may import and call those function.
    !
    ! Instructions:
    !   1. write a wrapper function that will be decorated (exported) that will wrap any 
    !      functionality (function calls / subroutine calls) required by this DLL
    !
    !   2. Imported the decorated functions from the DLL
    !
    !   3. Call the DLL functions through the wrapper
    !-----------------------------------------------------------------------------------------------

        !======================================================================
        ! HYCTL_EXPORTS
        !----------------------------------------------------------------------
        ! This module interfaces definitions of each of the functions to be
        ! imported and a copy of the handle of the DLL and all the function 
        ! pointers
        !----------------------------------------------------------------------
        MODULE HYCTL_EXPORTS            ! Module holds all of the pointers/interfaces for accessing our DLL
        USE, INTRINSIC::ISO_C_BINDING   ! provides access to required HANDLEs and Pointer types
        IMPLICIT NONE                   ! Implicit none will make debugging easier, by preventing complication of variables/types we did not declare.

            !------------------------------------------------------------
            ! INTERFACEs
            !
            ! 
            !------------------------------------------------------------    
            ABSTRACT INTERFACE                                                   ! We define an abstract interface as there is no concrete function behind it untill we load the DLL
				SUBROUTINE hycon_pscad_cinterface(xdata, xin, xout, xvar, State) ! We must provide what the function would look like if calling normally
				DOUBLE PRECISION,		INTENT(IN)	:: xdata(20)
				DOUBLE PRECISION,		INTENT(IN)	:: xin(100)
				DOUBLE PRECISION,		INTENT(OUT)	:: xout(50)
				DOUBLE PRECISION,		INTENT(OUT)	:: xvar(1)
				INTEGER,			INTENT(INOUT)	:: STATE
				END SUBROUTINE hycon_pscad_cinterface                            ! End of the subroutine definition
            END INTERFACE                                                        ! End of the interface

            !------------------------------------------------------------
            ! Function Pointers
            !
            ! This will require the list of function pointers that point
            ! to the function imported from the DLL
            !
            ! In this example we are importing Dll Function
            !------------------------------------------------------------   
            PROCEDURE(hycon_pscad_cinterface), POINTER :: HyCtlPointer => NULL() ! Define pointer to a function that matches the previously defined interface

            !------------------------------------------------------------
            ! DLL Handler
            !
            ! This is the handle pointer to the DLL
            !------------------------------------------------------------  
            INTEGER(C_INTPTR_T) :: DLL_HANDLE = 0                                ! This variable will hold the handle to the DLL, that we can can see if the DLL is loaded.

        END MODULE HYCTL_EXPORTS                                                 ! We have finished defining everything that we know about the DLL, we can end the module definition.

        !======================================================================
        ! HYCTL_IMPORT_ROUTINES
        !----------------------------------------------------------------------
        ! This function will import the DLL. 
        ! Note: this is standard DLL importing. Not specific to this example
        !----------------------------------------------------------------------
        SUBROUTINE HYCTL_IMPORT_ROUTINES        ! We need to define a routing that will load the DLL and extract the functions.

! Required for library import functions for Intel
! This will provided the LoadLibrary and GetProcAddress function automatically for Intel Fortran
!
#ifdef __INTEL_COMPILER
        USE KERNEL32
#endif     
        USE HYCTL_EXPORTS                   ! Required for DLL function definitions, and function pointers
        USE, INTRINSIC :: ISO_C_BINDING     ! Required for ISO Binding definitions
        IMPLICIT NONE 
  
! GFortran Requires using Interfaces to import the LoadLibrary and GetProcAddress
! Functions.
!
#ifdef __GFORTRAN__
        INTERFACE 
            FUNCTION LoadLibrary(lpFileName) BIND(C,NAME='LoadLibraryA')
            USE, INTRINSIC :: ISO_C_BINDING
            IMPLICIT NONE 
            CHARACTER(KIND=C_CHAR) :: lpFileName(*) 
            !GCC$ ATTRIBUTES STDCALL :: LoadLibrary 
            INTEGER(C_INTPTR_T) :: LoadLibrary 
            END FUNCTION LoadLibrary 

            FUNCTION GetProcAddress(hModule, lpProcName)  &
            BIND(C, NAME='GetProcAddress')
            USE, INTRINSIC :: ISO_C_BINDING
            IMPLICIT NONE
            !GCC$ ATTRIBUTES STDCALL :: GetProcAddress
            TYPE(C_FUNPTR) :: GetProcAddress
            INTEGER(C_INTPTR_T), VALUE :: hModule
            CHARACTER(KIND=C_CHAR) :: lpProcName(*)
            END FUNCTION GetProcAddress      
        END INTERFACE
#endif
   
            ! Import the DLL, We provide a relative path to the DLL that is going to be loaded. The path
            ! is relative to where the EXE will reside.
            !
            DLL_HANDLE = LoadLibrary(C_CHAR_'SMAHYC_022011R04.dll'//C_NULL_CHAR)
            
            ! Extract the function pointers by name. Each of the function is extracted using the GetProcAddress
            ! then converted to the correct Function Pointer type. The conversion of function pointers is
            ! compiler dependent so we must use the correct function to perform it by detecting which type of
            ! compiler is being used.
            !
#if defined (__INTEL_COMPILER)
            CALL C_F_POINTER(TRANSFER(GetProcAddress(DLL_HANDLE, C_CHAR_'hycon_pscad_cinterface'//C_NULL_CHAR), C_NULL_PTR), HyCtlPointer)
#elif defined (__GFORTRAN__)
            CALL C_F_PROCPOINTER(GetProcAddress(DLL_HANDLE, C_CHAR_'hycon_pscad_cinterface'//C_NULL_CHAR), HyCtlPointer)
#endif
     
        ENDSUBROUTINE ! After Extracting the function the subroutine is done.

        !======================================================================
        ! SCKLIB_WRAPPER
        !----------------------------------------------------------------------
        ! This is a function that the Component will call directly, it will
        ! pass the execution off the imported DLL
        !----------------------------------------------------------------------

        SUBROUTINE HyCon_PSCAD_FInterface(xdata, xin, xout, xvar, State)  
        USE HYCTL_EXPORTS ! Required for access to the DLL functions
            DOUBLE PRECISION,		INTENT(IN)	:: xdata(20)
            DOUBLE PRECISION,		INTENT(IN)	:: xin(100)
            DOUBLE PRECISION,		INTENT(OUT)	:: xout(50)
            DOUBLE PRECISION,		INTENT(OUT)	:: xvar(1)
            INTEGER,			INTENT(INOUT)	:: STATE
            ! Ensure that the DLL is loaded, if not, load the dll
            !
            IF ( DLL_HANDLE .EQ. 0 ) THEN
                 CALL HYCTL_IMPORT_ROUTINES()
            ENDIF
            ! Call the dll Function
            !
            CALL HyCtlPointer(xdata, xin, xout, xvar, STATE)

        END SUBROUTINE
















