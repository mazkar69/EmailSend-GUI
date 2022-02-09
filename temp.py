from traceback import print_tb
from urllib3 import Retry


def hcf(num1,num2,num3):
    hcf_=1
    def find_smallest(a,b,c): 
        if a<b and a<c:
            return a;
        elif b<a and b<c:
            return b;
        else:
            return c;

    snum=find_smallest(num1,num2,num3) 
    for i in range(2,snum+1):
        if(num1 % i == 0 and num2 % i == 0 and num3 % i ==0 ):
            hcf_=i;

        
    return hcf_




print(hcf(6,72,120))








