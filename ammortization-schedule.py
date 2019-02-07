import xlsxwriter
workbook= xlsxwriter.Workbook('Amortizationschedule.xlsx')
worksheet = workbook.add_worksheet()


class Mortgage:

    def __init__(self, amount, rate, years):
        self.amount = amount
        self.rate = rate
        self.years = years
        self.term = years*12
    
    def __repr(self):
        return(self.amount)
        
    ## calculates the monthly payment for the mortgage
    def pmt(self):
        rate = self.rate/2
        n = self.term
        i = ((1+rate)**(1/6))-1
        a = (1-(1+i)**(-n))/i
        r = self.amount/a
        return round(r,2)
    
    
    ## calculates the value of the last payment
    def lastpmt(self):
        rate = self.rate/2
        i = ((1+rate)**(1/6))-1
        n = self.term
        So = self.amount*(1+i)**n
        r = self.pmt()
        Sr = (r*((1+i)**(n-1)-1)/i)*(1+i)
        return round(So-Sr,2)
    
    ## calculates the value of the outstanding balance at time k(in months)
    def outbalance(self,k):
        rate = self.rate/2
        i = ((1+rate)**(1/6))-1
        r = self.pmt()
        Sr = r*((1+i)**(k)-1)/i
        Ak = self.amount*(1+i)**k
        return round(Ak-Sr,2)
    
    ## calculates how much of the monthly payment is devoted to the interest
    ## at time k
    def interestpmt(self,k):
        rate = self.rate/2
        i = ((1+rate)**(1/6))-1
        Ob = self.outbalance(k-1)
        return round(Ob*i,2)
        
    
    ## calculates how much of the monthly payment is devoted to the principal
    def principalpmt(self,k):
        p = self.pmt()
        ip = self.interestpmt(k)
        return round(p-ip,2)

   

    def maker(self):
        rate = self.rate/2
        i = ((1+rate)**(1/6))-1
        pmt = self.pmt()
        term = self.term
        lastpmt = self.lastpmt()
        lastprincipal = lastpmt-self.interestpmt(term)
        lastbalance=0
        
        
        for i in range(0,term+1):
            worksheet.write(i,0,i)
        
        for i in range(0,term):
            worksheet.write(i,1,pmt)
        
        for i in range(0,term+1):
            worksheet.write(i,2,self.interestpmt(i))
        
        for i in range(0,term):
            worksheet.write(i,3,self.principalpmt(i))
            
        for i in range(0,term):
            worksheet.write(i,4,self.outbalance(i))        
        
        worksheet.write(term,1,self.lastpmt())
        worksheet.write(term,3,lastprincipal)
        worksheet.write(term,4,lastbalance)
        
        worksheet.write(term+1,0,"Total")
        worksheet.write(term+1,1,'=Sum(B1:B301)')
        worksheet.write(term+1,2,'=Sum(C1:C301)')
        worksheet.write(term+1,3,'=Sum(D1:D301)')
        workbook.close()
        

    
        
        
        
        




   
    
        
        
        
        
        
