# DataRetriever_StockPrice

### Description
It scrap the stock price from yahoo financials using VBA

## Tech Stack
- **Frontend:** 
- **Backend:** VBA
- **Database:** 
- **Hardware:** 
- **Other Tools:**  

## Features
- It scrap the stock price from yahoo financials using VBA

### Installation

VBA

VBA code to get the stock price from yahoo financials

Below is the VBA code for the stock price retriever. 

Sub stockprice()
    Dim ie As InternetExplorer
    Dim strURL As String
    Dim i As Integer
    
    For i = 2 To Range("B1000").End(xlUp).Row
'    We select the second row as the beginning row and find the ending row within 1000 rows. End(xlup) enables us to find the last row that has data.
    
        strURL = "https://finance.yahoo.com/quote/" & Range("B" & i)
'       We are going to get the stock price from Yahoo Financial
        
        Set ie = CreateObject("InternetExplorer.application")
        
        ie.Navigate strURL
        ie.Visible = True
        
        Do While (ie.ReadyState <> READYSTATE_COMPLETE Or ie.Busy = True)
            DoEvents
        Loop
        
        Range("c" & i) = ie.Document.getelementsbyclassname("Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)")(0).innertext
'        The stock price is in the above class, so we search for the stock price with the above class name.

        ie.Quit
        Set ie = Nothing
    Next i

    
End Sub
