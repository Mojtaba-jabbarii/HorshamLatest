! ******************************************************
!             Subroutine to Interface PSCAD with SMA Simulink Models
! ******************************************************
      SUBROUTINE SMA_SCK_FInterface(xdata, xin, xout,&
     & UpDown, DclVolSpt,&
     & PCSFb, fdebugout)    

      INCLUDE "nd.h"         ! dimensions          
                             ! in the case of EMTDC with Digital Fortran 90
                             ! this file introduces the module NDDE with a array
                             ! E_NDIM which contains all the parameters
                             ! defining sizes of EMTDC matrixes and
                             ! arrays
 
      INCLUDE "s1.h"         ! time etc...
      INCLUDE "s2.h"         ! array STOR and index NEXC
                             ! in the case of EMTDC with Digital Fortarn 90
                             ! arrays STOR is dimentioning in the follwing way
                             !
                             ! real - STOR(E_NDIM(10))

      INCLUDE "emtstor.h"    ! arrays and indexes: STORx, NSTORx
                             ! in the case of EMTDC with Digital Fortarn 90
                             ! arrays STORx are dimentioning in the follwing way
                             ! 
                             ! logical - STORL(E_NIDM(27))  
                             ! integer - STORI(E_NDIM(28))
                             ! real    - STORF(E_NDIM(29))
                             ! complex - STORC(E_NDIM(30))


      DOUBLE PRECISION xdata(10) 		! config data
      DOUBLE PRECISION xin(50) 			! inputs
      DOUBLE PRECISION xout(50)			! outputs


! Input Parameters
      INTEGER          UpDown
      DOUBLE PRECISION DclVolSpt
     

! Outputs
      DOUBLE PRECISION PCSFb(10)			! Output PSC feedbacks WRtgVArRtg, PQ Available
      DOUBLE PRECISION fdebugout(100)    			! Output Debug Real(100)
      



! Fortran 90 interface to a C procedure
      INTERFACE
         SUBROUTINE sma_sck_cinterface(xdata, xin, xout,&
     & UpDown,DclVolSpt,PCSFb,&
     & fdebugout, State) &
		 & bind (C, name="sma_sck_cinterface") 

            !DEC$ ATTRIBUTES REFERENCE :: xdata
            !DEC$ ATTRIBUTES REFERENCE :: xin
            !DEC$ ATTRIBUTES REFERENCE :: xout

            !DEC$ ATTRIBUTES REFERENCE :: UpDown

            !DEC$ ATTRIBUTES REFERENCE :: DclVolSpt

            !DEC$ ATTRIBUTES REFERENCE :: PCSFb
            !DEC$ ATTRIBUTES REFERENCE :: fdebugout
            !DEC$ ATTRIBUTES REFERENCE :: State

! Input Variables
            DOUBLE PRECISION xdata(10)
            DOUBLE PRECISION xin(50)
            DOUBLE PRECISION xout(50)

! Input Parameters
            INTEGER          UpDown
            DOUBLE PRECISION DclVolSpt
! Output Variables
            DOUBLE PRECISION PCSFb(10)
            DOUBLE PRECISION fdebugout(100)
            INTEGER          State(20000)
         END SUBROUTINE
     END INTERFACE



      CALL sma_sck_cinterface(xdata, xin, xout, &
     & UpDown, DclVolSpt,PCSFb,&
     & fdebugout, STORI(NSTORI))


! Note the storage pointer is incremented in the main component
! This is done because the model is not called every time step, but the pointer
! must be incremented each step...
      RETURN
      END


