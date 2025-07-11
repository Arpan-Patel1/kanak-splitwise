,sequence,column_name,field_level,type,aliases,sample,description,controls_definitions
0,6,Product,table,String,,Securities,"The financial product being offered. If no product is specified, it should be ""Cash"". Do not leave this field empty. ","Should be one of ""Equities"", ""Corporate Fixed Income"", ""Government Fixed Income"", ""Money Market"", ""Cash"", or ""Derivative"" ."
1,7,Currency,table,String,,USD,The currency.,"Either the currency or the country fields should not be blank or null, else it is not a valid row or a table."
2,8,Country,table,String,,CA,The country for the instructions.,"Either the currency or the country fields should not be blank or null, else it is not a valid row or a table."
3,9,Agent Bank Name,table,String,,"['Royal Bank of Canada, Toronto.', 'Please see special instructions on final pages.']",The Name of the agent bank.,Should be max 35 characters.
4,10,Agent Bank BIC,table,String,['Agent BIC'],DEUTHKHHXXX,BIC code for the agent bank.,Should be 8 to 11 characters long with only letters and numbers.
5,11,Account or Beneficiary Bank Name,table,String,,NOMURA INTERNATIONAL PLC,The Name of the beneficiary bank.,Should be max 35 characters.
6,12,Agent Bank Account Number,table,String,,DK83 3000 3996 0176 72,The Account Number for the agent bank.,Should be max 35 characters.
7,13,IBAN,table,String,,,International Bank Account Number.,"Should be blank if product is not ""Cash"" and should be 22 non-whitespace characters."
8,14,PSET or CSD BIC,table,String,['CSD BIC'],,Place of Settlement BIC or the SWIFT CODE. ,
9,15,PSET or CSD Account Name,table,String,['CSD name'],,The central securities depository CSD name.,
10,16,PSET or CSD Account Number,table,String,['CSD Account No/Ref'],,The central securities depository account number.,"Should be a number or a alphanumeric code. e.g. RBCT,  4504. "
11,1,Document Name,document,String,,Standard Settlement Instructions,What is the name of the document? ,
12,2,Date,document,String,,6/15/2024,What is the main filing date for the document if available or the document date?,
13,3,Legal Entity,document,String,,Morgan Stanley & Co International PLC,What is the main legal entity associated with this document?,
14,4,LEI,document,String,,4PQUHN3JPFGFNF3BB653,The Legal Entity ID for the main legal entity associated with this document.,
15,5,Beneficiary Bank BIC,table,String,,MSLNGB2X,The BIC for the main Legal Entity of this document.,
16,17,US TradeSuite ID,table,String,,61974,The US TradeSuite ID.,
17,18,Ultimate Beneficiary Bank Name,table,String,,"Royal Bank of Canada, Toronto.","The ultimate beneficary bank. Appears only after ""For the account of"". ",Should be max 35 characters. Should be different than the Agent Bank Name and the Beneficiary Bank Name.







[
  {
    "Country": "EQUITY SETTLEMENT INSTRUCTIONS",
    "Currency": "",
    "Agent Bank Name (DEAG/REAG)": "",
    "Agent BIC CSD BIC (PSET)": "",
    "Account Name (BUYR/SELL)": "",
    "CSD Name (PSET)": "",
    "Account No/Ref CSD Acc No/Ref (PSET)": ""
  },
  {
    "Country": "AUSTRALIA",
    "Currency": "AUD",
    "Agent Bank Name (DEAG/REAG)": "CITIBANK LIMITED MELBOURNE",
    "Agent BIC CSD BIC (PSET)": "CITIAU3X CAETAU21",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC CLEARING HSE ELECTRONIC SUB SYS",
    "CSD Name (PSET)": "SYD",
    "Account No/Ref CSD Acc No/Ref (PSET)": "2010540000 20018"
  },
  {
    "Country": "AUSTRIA",
    "Currency": "EUR",
    "Agent Bank Name (DEAG/REAG)": "UNICREDIT BANK AUSTRIA AG WIE",
    "Agent BIC CSD BIC (PSET)": "BKAUATWW OEKOATWW",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC OESTERREICHISCHE KONTROLLBK AG",
    "CSD Name (PSET)": "-VIE",
    "Account No/Ref CSD Acc No/Ref (PSET)": "0101-08462 00 222100"
  },
  {
    "Country": "BAHRAIN",
    "Currency": "USD",
    "Agent Bank Name (DEAG/REAG)": "HSBC BANK MIDDLE EAST MANAMA",
    "Agent BIC CSD BIC (PSET)": "BBMEBHBX XBAHBHB1",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC BAHRIAN STOCK EXCHANGE",
    "CSD Name (PSET)": "MANAMA",
    "Account No/Ref CSD Acc No/Ref (PSET)": "001-005107-085"
  },
  {
    "Country": "BELGIUM",
    "Currency": "",
    "Agent Bank Name (DEAG/REAG)": "Please see special instructions on final pages",
    "Agent BIC CSD BIC (PSET)": "",
    "Account Name (BUYR/SELL)": "",
    "CSD Name (PSET)": "",
    "Account No/Ref CSD Acc No/Ref (PSET)": ""
  },
  {
    "Country": "BOTSWANA",
    "Currency": "BWP",
    "Agent Bank Name (DEAG/REAG)": "STD CHAT BK BOTSWANA LTD",
    "Agent BIC CSD BIC (PSET)": "SCHBBWGX XBOTBWG1",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC BOTSWANA STOCK",
    "CSD Name (PSET)": "EXCHANGE",
    "Account No/Ref CSD Acc No/Ref (PSET)": "0341"
  },
  {
    "Country": "BRAZIL",
    "Currency": "",
    "Agent Bank Name (DEAG/REAG)": "Please see special instructions on final pages",
    "Agent BIC CSD BIC (PSET)": "",
    "Account Name (BUYR/SELL)": "",
    "CSD Name (PSET)": "",
    "Account No/Ref CSD Acc No/Ref (PSET)": ""
  },
  {
    "Country": "BULGARIA",
    "Currency": "BGN",
    "Agent Bank Name (DEAG/REAG)": "STEP OUT MARKET BGL",
    "Agent BIC CSD BIC (PSET)": "",
    "Account Name (BUYR/SELL)": "STEP OUT MARKET",
    "CSD Name (PSET)": "BGL",
    "Account No/Ref CSD Acc No/Ref (PSET)": ""
  },
  {
    "Country": "CANADA",
    "Currency": "CAD",
    "Agent Bank Name (DEAG/REAG)": "ROYAL BANK OF CANADA TORONTO",
    "Agent BIC CSD BIC (PSET)": "ROYCCAT2 CDSLCATT",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC CANADIAN DEPOSITORY FOR",
    "CSD Name (PSET)": "SECURITIES",
    "Account No/Ref CSD Acc No/Ref (PSET)": "T12897981 RBCT"
  },
  {
    "Country": "CANADA",
    "Currency": "EUR",
    "Agent Bank Name (DEAG/REAG)": "ROYAL BANK OF CANADA TORONTO",
    "Agent BIC CSD BIC (PSET)": "ROYCCAT2 CDSLCATT",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC CANADIAN DEPOSITORY FOR",
    "CSD Name (PSET)": "SECURITIES",
    "Account No/Ref CSD Acc No/Ref (PSET)": "T12897981 RBCT"
  },
  {
    "Country": "CANADA",
    "Currency": "GBP",
    "Agent Bank Name (DEAG/REAG)": "ROYAL BANK OF CANADA TORONTO",
    "Agent BIC CSD BIC (PSET)": "ROYCCAT2 CDSLCATT",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC CANADIAN DEPOSITORY FOR",
    "CSD Name (PSET)": "SECURITIES",
    "Account No/Ref CSD Acc No/Ref (PSET)": "T12897981 RBCT"
  },
  {
    "Country": "CANADA",
    "Currency": "USD",
    "Agent Bank Name (DEAG/REAG)": "ROYAL BANK OF CANADA TORONTO",
    "Agent BIC CSD BIC (PSET)": "ROYCCAT2 CDSLCATT",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC CANADIAN DEPOSITORY FOR",
    "CSD Name (PSET)": "SECURITIES",
    "Account No/Ref CSD Acc No/Ref (PSET)": "T12897981 RBCT"
  },
  {
    "Country": "CHINA",
    "Currency": "",
    "Agent Bank Name (DEAG/REAG)": "Please see special instructions on final pages",
    "Agent BIC CSD BIC (PSET)": "",
    "Account Name (BUYR/SELL)": "",
    "CSD Name (PSET)": "",
    "Account No/Ref CSD Acc No/Ref (PSET)": ""
  },
  {
    "Country": "CROATIA",
    "Currency": "HRK",
    "Agent Bank Name (DEAG/REAG)": "ZAGREBACKA BANKA ZAGREB",
    "Agent BIC CSD BIC (PSET)": "ZABAHR2X SDAHHR22",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC SREDISNJA DEPOZIT AGENCIJA",
    "CSD Name (PSET)": "ZAGREB",
    "Account No/Ref CSD Acc No/Ref (PSET)": "9991950006488005999"
  },
  {
    "Country": "CZECH REPUBLIC",
    "Currency": "CZK",
    "Agent Bank Name (DEAG/REAG)": "UNICREDIT BANK CZECH REPUBLIC",
    "Agent BIC CSD BIC (PSET)": "BACXCZPP UNIYCZPP",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC CENTRALNI DEPOZITAR CENNYCH",
    "CSD Name (PSET)": "PAPIRU",
    "Account No/Ref CSD Acc No/Ref (PSET)": "81252000"
  },
  {
    "Country": "CZECH REPUBLIC",
    "Currency": "USD",
    "Agent Bank Name (DEAG/REAG)": "UNICREDIT BANK CZECH REPUBLIC",
    "Agent BIC CSD BIC (PSET)": "BACXCZPP UNIYCZPP",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC CENTRALNI DEPOZITAR CENNYCH",
    "CSD Name (PSET)": "PAPIRU",
    "Account No/Ref CSD Acc No/Ref (PSET)": "81252000"
  },
  {
    "Country": "DENMARK",
    "Currency": "",
    "Agent Bank Name (DEAG/REAG)": "Please see special instructions on final pages",
    "Agent BIC CSD BIC (PSET)": "",
    "Account Name (BUYR/SELL)": "",
    "CSD Name (PSET)": "",
    "Account No/Ref CSD Acc No/Ref (PSET)": ""
  },
  {
    "Country": "EGYPT",
    "Currency": "EGP",
    "Agent Bank Name (DEAG/REAG)": "CITIBANK CAIRO NOSTRO ACCOUNT",
    "Agent BIC CSD BIC (PSET)": "CITIEGCX MCSDEGCA",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC MCSD MISR FOR CLRNG SMENT+DEP",
    "CSD Name (PSET)": "CAIRO",
    "Account No/Ref CSD Acc No/Ref (PSET)": "1350031300 4504"
  },
  {
    "Country": "ESTONIA",
    "Currency": "EEK",
    "Agent Bank Name (DEAG/REAG)": "SWEDBANK ESTONIA",
    "Agent BIC CSD BIC (PSET)": "HABAEE2X",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC ESTONIAN CENTRAL DEP FOR",
    "CSD Name (PSET)": "SECS",
    "Account No/Ref CSD Acc No/Ref (PSET)": "99000539922"
  },
  {
    "Country": "ESTONIA",
    "Currency": "EUR",
    "Agent Bank Name (DEAG/REAG)": "SWEDBANK ESTONIA",
    "Agent BIC CSD BIC (PSET)": "HABAEE2X",
    "Account Name (BUYR/SELL)": "NOMURA INTERNATIONAL PLC ESTONIAN CENTRAL DEP FOR",
    "CSD Name (PSET)": "SECS",
    "Account No/Ref CSD Acc No/Ref (PSET)": "99000539922"
  },
  {
    "Country": "FINLAND",
    "Currency": "",
    "Agent Bank Name (DEAG/REAG)": "Please see special instructions on final pages",
    "Agent BIC CSD BIC (PSET)": "",
    "Account Name (BUYR/SELL)": "",
    "CSD Name (PSET)": "",
    "Account No/Ref CSD Acc No/Ref (PSET)": ""
  }
]
                                                                                                               







[
  {
    "instrumentType": "EQUITY",
    "country": null,
    "settlementCurrency": null,
    "settlementMethod": null,
    "settlementDate": null,
    "settlementAmount": null,
    "counterpartyName": null,
    "counterpartyAccount": null,
    "instructions": null
  },
  {
    "country": "AUSTRALIA",
    "currency": "AUD",
    "agentBankName": "CITIBANK LIMITED MELBOURNE",
    "agentBic": "CITIAU3X",
    "csdBic": "CAETAU21",
    "accountName": "NOMURA INTERNATIONAL PLC CLEARING HSE ELECTRONIC SUB SYS",
    "csdName": "SYD",
    "accountNumber": "2010540000",
    "csdAccountNumber": "20018"
  },
  {
    "country": "AUSTRIA",
    "currency": "EUR",
    "agentBankName": "UNICREDIT BANK AUSTRIA AG WIE",
    "agentBic": "BKAUATWW",
    "csdBic": "OEKOATWW",
    "accountName": "NOMURA INTERNATIONAL PLC OESTERREICHISCHE KONTROLLBK AG",
    "csdName": "-VIE",
    "accountNumber": "0101-08462",
    "accountReference": "00 222100"
  },
  {
    "country": "BAHRAIN",
    "currency": "USD",
    "agentBankName": "HSBC BANK MIDDLE EAST MANAMA",
    "agentBic": "BBMEBHBX",
    "csdBic": "XBAHBHB1",
    "accountName": "NOMURA INTERNATIONAL PLC BAHRIAN STOCK EXCHANGE",
    "csdName": "MANAMA",
    "accountNumber": "001-005107-085"
  },
  {
    "country": "BELGIUM",
    "agentBankName": null
  },
  {
    "country": "BOTSWANA",
    "currency": "BWP",
    "agentBankName": "STD CHAT BK BOTSWANA LTD",
    "agentBic": "SCHBBWGX",
    "accountName": "NOMURA INTERNATIONAL PLC BOTSWANA STOCK",
    "csdName": "EXCHANGE",
    "accountNumber": "0341"
  },
  {
    "country": "BRAZIL",
    "agentBankName": null
  },
  {
    "country": "BULGARIA",
    "currency": "BGN",
    "agentBankName": "STEP OUT MARKET BGL",
    "accountName": "STEP OUT MARKET",
    "csdName": "BGL"
  },
  {
    "country": "CANADA",
    "currency": "CAD",
    "agentBankName": "ROYAL BANK OF CANADA TORONTO",
    "agentBic": "ROYCCAT2",
    "accountName": "NOMURA INTERNATIONAL PLC CANADIAN DEPOSITORY FOR",
    "csdName": "SECURITIES",
    "accountNumber": "T12897981",
    "csdAccountNumber": "RBCT"
  },
  {
    "country": "CANADA",
    "currency": "EUR",
    "agentBankName": "ROYAL BANK OF CANADA TORONTO",
    "agentBic": "ROYCCAT2",
    "accountName": "NOMURA INTERNATIONAL PLC CANADIAN DEPOSITORY FOR",
    "csdName": "SECURITIES",
    "accountNumber": "T12897981",
    "csdAccountNumber": "RBCT"
  },
  {
    "country": "CANADA",
    "currency": "GBP",
    "agentBankName": "ROYAL BANK OF CANADA TORONTO",
    "agentBic": "ROYCCAT2",
    "accountName": "NOMURA INTERNATIONAL PLC CANADIAN DEPOSITORY FOR",
    "csdName": "SECURITIES",
    "accountNumber": "T12897981",
    "csdAccountNumber": "RBCT"
  },
  {
    "country": "CANADA",
    "currency": "USD",
    "agentBankName": "ROYAL BANK OF CANADA TORONTO",
    "agentBic": "ROYCCAT2",
    "accountName": "NOMURA INTERNATIONAL PLC CANADIAN DEPOSITORY FOR",
    "csdName": "SECURITIES",
    "accountNumber": "T12897981",
    "csdAccountNumber": "RBCT"
  },
  {
    "country": "CHINA",
    "agentBankName": null
  },
  {
    "country": "CROATIA",
    "currency": "HRK",
    "agentBankName": "ZAGREBACKA BANKA ZAGREB",
    "agentBic": "ZABAHR2X",
    "csdBic": "SDAHHR22",
    "accountName": "NOMURA INTERNATIONAL PLC SREDISNJA DEPOZIT AGENCIJA",
    "csdName": "ZAGREB",
    "accountNumber": "9991950006488005999"
  },
  {
    "country": "CZECH REPUBLIC",
    "currency": "CZK",
    "agentBankName": "UNICREDIT BANK CZECH REPUBLIC",
    "agentBic": "BACXCZPP",
    "csdBic": "UNIYCZPP",
    "accountName": "NOMURA INTERNATIONAL PLC CENTRALNI DEPOZITAR CENNYCH",
    "csdName": "PAPIRU",
    "accountNumber": "81252000"
  },
  {
    "country": "CZECH REPUBLIC",
    "currency": "USD",
    "agentBankName": "UNICREDIT BANK CZECH REPUBLIC",
    "agentBic": "BACXCZPP",
    "csdBic": "UNIYCZPP",
    "accountName": "NOMURA INTERNATIONAL PLC CENTRALNI DEPOZITAR CENNYCH",
    "csdName": "PAPIRU",
    "accountNumber": "81252000"
  },
  {
    "country": "DENMARK",
    "agentBankName": null
  },
  {
    "country": "EGYPT",
    "currency": "EGP",
    "agentBankName": "CITIBANK CAIRO NOSTRO ACCOUNT",
    "agentBic": "CITIEGCX",
    "csdBic": "MCSDEGCA",
    "accountName": "NOMURA INTERNATIONAL PLC MCSD MISR FOR CLRNG SMENT+DEP",
    "csdName": "CAIRO",
    "accountNumber": "1350031300",
    "accountReference": "4504"
  },
  {
    "country": "ESTONIA",
    "currency": "EEK",
    "agentBankName": "SWEDBANK ESTONIA",
    "agentBic": "HABAEE2X",
    "accountName": "NOMURA INTERNATIONAL PLC ESTONIAN CENTRAL DEP FOR",
    "csdName": "SECS",
    "accountNumber": "99000539922"
  },
  {
    "country": "ESTONIA",
    "currency": "EUR",
    "agentBankName": "SWEDBANK ESTONIA",
    "agentBic": "HABAEE2X",
    "accountName": "NOMURA INTERNATIONAL PLC ESTONIAN CENTRAL DEP FOR",
    "csdName": "SECS",
    "accountNumber": "99000539922"
  },
  {
    "country": "FINLAND",
    "agentBankName": null
  }
]
