# -*- coding: utf-8 -*-
"""
Created on Wed May 20 14:02:54 2020

@author: Mervin Kall
"""
import cmath
import math

#returns Grid impedance in pu (X and R) expressed on Vbase, Sbase_Z.
def calc_Z(SCR, X_R, Sbase_plant,  Sbase_Z, Vbase):
    SscGrid=Sbase_plant*SCR*1000000 #Grid SSC in VA
    Zgrid = (Vbase*Vbase)/SscGrid #Zgrid in Ohms
    
    #calculate R and X in Ohms
    R=Zgrid/math.sqrt((1+X_R*X_R))
    X=R*X_R
    
    #convert to pu
    Zbase=Vbase*Vbase/Sbase_Z
    
    return R/Zbase, X/Zbase #returns impedances in pu, on Vbase, Sbase_Z   


#returns the thevenin 
def calc_Vth_pu(X_R, SCC, Vbase, Qpoc, Ppoc, Vpoc):
    
#    X_R=10.0 #
#    SCR=3.618 #
#    Sbase=83 #Sbase of plant in MVA
#    Vbase=132000 #Vbase in V
#    
#    Qpoc=0.0 #Qpoc in MVAr
#    
#    Ppoc=83 #Ppoc in MW
#    
#    Vpoc=1.04 # Upoc in p.u.
    
       
    pi=3.1415926535897932384626433
    
    f=50.0
    
    Upoc=Vpoc*Vbase
    SscGrid=SCC*1000000
    
    
    Zgrid=(Vbase*Vbase)/SscGrid
    
    R=Zgrid/math.sqrt((1+X_R*X_R))
    X=R*X_R*1j
    
    ZgridC=R+X
    
    #R=Z/(1+X_R^2)
    #
    #2*pi*f*1j*L
    
    S=Qpoc*1000000*1j+Ppoc*1000000
    
    
    Ipoc=S/( (math.sqrt(3)) * Vbase*Vpoc) 
    Ipoc=Ipoc.conjugate()
    
    U1=ZgridC*Ipoc*math.sqrt(3)
    
    Uth=Upoc-U1
    
    Vth_pu=abs(Uth)/Vbase
    angle=cmath.phase(Uth)/3.14159265358979*180 #return angle of Vth in deg.
    
    print(Vth_pu)
    
    return Vth_pu, angle


def main():
#    calc_Vth_pu(X_R=3.07,SCR=7.5,Sbase=31.168, Vbase=132000, Qpoc=-3, Ppoc=25, Vpoc=1.04)
    calc_Vth_pu(X_R=3.07,SCC=7.5, Vbase=132000, Qpoc=-3, Ppoc=25, Vpoc=1.04)
    
if __name__ == '__main__':
    main()
XR=4.4
SC=557
#calc_Vth_pu(X_R=XR,SCC=SC, Vbase=220000, Qpoc=0, Ppoc=70, Vpoc=1.02)
#calc_Vth_pu(X_R=XR,SCC=SC, Vbase=220000, Qpoc=42, Ppoc=140, Vpoc=1.02)
#calc_Vth_pu(X_R=XR,SCC=SC, Vbase=220000, Qpoc=-42, Ppoc=140, Vpoc=1.02)
#calc_Vth_pu(X_R=XR,SCC=SC, Vbase=220000, Qpoc=-0, Ppoc=7, Vpoc=1.02)
#calc_Vth_pu(X_R=XR,SCC=SC, Vbase=220000, Qpoc=42, Ppoc=7, Vpoc=1.02)
calc_Vth_pu(X_R=XR,SCC=SC, Vbase=66000, Qpoc=0, Ppoc=80, Vpoc=0.97)

