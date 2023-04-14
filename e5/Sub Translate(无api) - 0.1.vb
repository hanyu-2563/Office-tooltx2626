Sub Translate()
    ' Declare variables
    Dim sourceText As String
    Dim targetText As String
    Dim sourceLanguage As String
    Dim targetLanguage As String
    Dim apiKey As String
    Dim apiUrl As String
    Dim response As String
    
    ' Get the selected text from the document
    sourceText = Selection.Text
    
    ' Check if the text is not empty
    If sourceText <> "" Then
        
        ' Set the source and target languages (use language codes)
        sourceLanguage = "en"
        targetLanguage = "zh"
        
        ' Set the API key and URL (use your own values)
        apiKey = "your_api_key"
        apiUrl = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&from=" & sourceLanguage & "&to=" & targetLanguage
        
        ' Create an HTTP request object
        Dim httpRequest As Object
        Set httpRequest = CreateObject("MSXML2.XMLHTTP")
        
        ' Set the request headers
        httpRequest.Open "POST", apiUrl, False
        httpRequest.setRequestHeader "Content-Type", "application/json"
        httpRequest.setRequestHeader "Ocp-Apim-Subscription-Key", apiKey
        
        ' Set the request body (use JSON format)
        Dim requestBody As String
        requestBody = "[{""Text"":""" & sourceText & """}]"
        
        ' Send the request
        httpRequest.send requestBody
        
        ' Get the response
        response = httpRequest.responseText
        
        ' Parse the response (use JSON format)
        Dim jsonResponse As Object
        Set jsonResponse = JsonConverter.ParseJson(response)
        
        ' Get the translated text from the response
        targetText = jsonResponse(1)("translations")(1)("text")
        
        ' Insert the translated text into the document
        Selection.TypeText targetText
        
    Else
        
        ' Display a message if the text is empty
        MsgBox "Please select some text to translate."
        
    End If
    
End Sub