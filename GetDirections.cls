'Edited to use positionstack free API instead of google maps

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub GetDirections()
    Dim googlekey As String
    Dim origin As String
    Dim destination As String
    Dim query As String
    Dim directions As String
    Dim xhrRequest As Object
    Dim domDoc As Object
    
    'Set the variables
    googlekey = "<49a24b3bf1d06e7ae131888ff5d6b3f8>"
    origin = Replace(Range("A2"), ", ", "+")
    destination = Replace(Range("A5"), ", ", "+")
    query = "origin=" & origin & "&destination=" & destination
    
    'Create the objects
    Set xhrRequest = CreateObject("MSXML2.serverXMLHTTP")
    Set domDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    'Clear the directions cells
    Range("A11:A100").Value = ""
    
    'Check to make sure we have input in both origin and destination
    If Range("A2") = "" Then
        MsgBox ("Need to input Origin")
    ElseIf Range("A5") = "" Then
        MsgBox ("Need to input Province/State")
    Else
        'Open and send the request for directions to Google Maps API
        xhrRequest.Open "GET", "http://api.positionstack.com/v1/" & _
        query & googlekey, False
        xhrRequest.Send
    
      'will output as json by default - can output as xml on paid plan, which I don't want, but I don't know how to parse json because I have no clue what I'm doing
      
        'Load the XML into a container
        domDoc.LoadXML xhrRequest.responseText
        
        'Get all directions
        Set Nodes = domDoc.getElementsByTagName("html_instructions")
        
        'Loop through them and output them to the sheet
        Count = 11
        For Each Node In Nodes
            c = "A" & Count
            directions = Node.Text
            Range(c).Value = directions
            'Remove html tags
            With CreateObject("vbscript.regexp")
                .Pattern = "\<.*?\>"
                .Global = True
                Range(c).Value = .Replace(Range(c).Value, "")
            End With
            Count = Count + 1
        Next Node
    End If
End Sub
