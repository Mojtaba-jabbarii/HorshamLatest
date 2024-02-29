! ******************************************************
!  Subroutine to Interface PSCAD with SMA HyCon model
! ******************************************************
      SUBROUTINE HyCon_PSCAD_FInterface(xdata, xin, xout, xvar)    

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



      DOUBLE PRECISION xdata(20) 	! config data
      DOUBLE PRECISION xin(100) 	! inputs
      DOUBLE PRECISION xout(40) 	! outputs
      DOUBLE PRECISION xvar(1) 		! variables      



! Fortran 90 interface to a C procedure
      INTERFACE
         SUBROUTINE hycon_pscad_cinterface(xdata, xin, xout, xvar, State) &
		 & bind (C, name="hycon_pscad_cinterface") 

            !DEC$ ATTRIBUTES REFERENCE :: xdata
            !DEC$ ATTRIBUTES REFERENCE :: xin
            !DEC$ ATTRIBUTES REFERENCE :: xout
            !DEC$ ATTRIBUTES REFERENCE :: xvar
            !DEC$ ATTRIBUTES REFERENCE :: State

! Input Variables
            DOUBLE PRECISION xdata(20)
            DOUBLE PRECISION xin(100)
            DOUBLE PRECISION xout(40)
            DOUBLE PRECISION xvar(1)
            INTEGER          State(50000)

         END SUBROUTINE
     END INTERFACE



      CALL hycon_pscad_cinterface(xdata, xin, xout, xvar,STORI(NSTORI))

! Note the storage pointer is incremented in the main component
! This is done because the model is not called every time step, but the pointer
! must be incremented each step...
      RETURN
      END


