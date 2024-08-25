Attribute VB_Name = "MergeData"
'*********************************************************
'Created by : Leonardo Sembiring
'Date       : 23 June 2024
'Purpose    : To merge data from Excel to Word document
'*********************************************************

Public Sub MergeData()
    m_sheetgeneral = "General" 'main sheet of this file
    
    m_foldername = Sheets(m_sheetgeneral).Cells(2, 2).Value 'working folder
    m_sourceworkbook = Sheets(m_sheetgeneral).Cells(5, 2).Value 'workbook name
    m_sourcesheet = Sheets(m_sheetgeneral).Cells(6, 2).Value 'worksheet name of source
    m_destination = Sheets(m_sheetgeneral).Cells(18, 2).Value 'destination word file
    
    'open file word : destination
    Set WordApp = CreateObject("word.Application")
    WordApp.Documents.Open m_foldername & "\" & m_destination
    WordApp.Visible = True
    Set WordDocument = WordApp.Documents(m_foldername & "\" & m_destination)
    
    m_totalnilaitagih = 0
    m_policynumberstring = ""
    
    m_row = 2
    m_keygrouping = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 7).Value
    m_firstpolicy = True
    
    Do While Not IsEmpty(Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 1).Value)
        '**************************************************************************
        'evaluates policy information
        m_policynumber = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 1).Value
        m_productname = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 2).Value
        m_status = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 3).Value
        m_dasar = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 4).Value
        m_nilaitagih = Val(Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 5).Value)
        m_nama = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 6).Value
        m_pempolid = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 7).Value
        m_nowa = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 8).Value
        m_norek = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 9).Value
        m_namabank = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 10).Value
        m_namarek = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row, 11).Value
        '**************************************************************************
        
        If m_firstpolicy Then
            m_tablerow = 1
            '**************************************************************************
            'write here for policyholder information via "bookmark" in Word
            WordDocument.Bookmarks("pempolid").Range.Text = m_pempolid
            WordDocument.Bookmarks("nama").Range.Text = m_nama
            WordDocument.Bookmarks("nama1").Range.Text = m_nama
            WordDocument.Bookmarks("nowa").Range.Text = m_nowa
            WordDocument.Bookmarks("norek").Range.Text = m_norek
            WordDocument.Bookmarks("namarek").Range.Text = m_namarek
            WordDocument.Bookmarks("namabank").Range.Text = m_namabank
            '**************************************************************************
            m_firstpolicy = False
        End If
       
        'recording total nilai tagih
        m_totalnilaitagih = m_totalnilaitagih + m_nilaitagih
        m_policynumberstring = m_policynumberstring & m_policynumber & ","
        
        'added to the table : list of policies
        WordDocument.Tables(1).Rows.Add
        m_tablerow = m_tablerow + 1
        
        '**************************************************************************
        'populate the table
        WordDocument.Tables(1).Cell(m_tablerow, 1).Range.Text = m_tablerow - 1
        WordDocument.Tables(1).Cell(m_tablerow, 2).Range.Text = m_policynumber
        WordDocument.Tables(1).Cell(m_tablerow, 3).Range.Text = m_productname
        WordDocument.Tables(1).Cell(m_tablerow, 4).Range.Text = m_status
        WordDocument.Tables(1).Cell(m_tablerow, 5).Range.Text = m_dasar
        WordDocument.Tables(1).Cell(m_tablerow, 6).Range.Text = Format(m_nilaitagih, "##,###,###,###")
        '**************************************************************************
        
        If Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row + 1, 7).Value _
            <> m_keygrouping Then
            
            '**************************************************************************
            'write the total nilai tagih
            WordDocument.Tables(1).Cell(m_tablerow + 1, 1).Range.Text = "Total"
            WordDocument.Tables(1).Cell(m_tablerow + 1, 6).Range.Text = Format(m_totalnilaitagih, "##,###,###,###")
            'merge row total
            With WordDocument.Tables(1)
                .Cell(Row:=m_tablerow + 1, Column:=1).Merge _
                MergeTo:=.Cell(Row:=m_tablerow + 1, Column:=5)
            End With
            'font bold for total
            WordDocument.Tables(1).Rows(m_tablerow + 1).Range.Select
            WordApp.Selection.Font.Bold = True
            '**************************************************************************
            
            'save as word document : parsing ".docx"
            WordDocument.SaveAs (m_foldername & "\" & "SPHT-" & m_pempolid & ".docx")
            WordDocument.Close
            
            'next policyholder name
            m_totalnilaitagih = 0
            m_keygrouping = Workbooks(m_sourceworkbook).Sheets(m_sourcesheet).Cells(m_row + 1, 7).Value
            m_firstpolicy = True
            
            'reopen the word template
            WordApp.Documents.Open m_foldername & "\" & m_destination
            WordApp.Visible = True
            Set WordDocument = WordApp.Documents(m_foldername & "\" & m_destination)
    
        End If
        m_row = m_row + 1
    Loop
    WordDocument.Close
End Sub




