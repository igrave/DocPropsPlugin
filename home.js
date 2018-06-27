(function () {
    "use strict";
 
    // The initialize function is run each time the page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
 
            // Use this to check whether the API is supported in the Word client.
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                // Do something that is only available via the new APIs
             //   $('#emerson').click(insertEmersonQuoteAtSelection);
             //   $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
             //   $('#proverb').click(insertChineseProverbAtTheEnd);
             //   $('#fieldadd').click(insertfieldxml);
                $('#fieldaddname').click(insertfieldxmlname);
                $('#props').click(showProps);
                $('#supportedVersion').html('This code is using Word 2016 or greater.');
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or greater.');
            }
        });
    };
    
    
    function showProps(){
      Word.run(function(context){
        var properties = context.document.properties;
        $('#wordProps').html("Hello!");
        context.load(properties);
        var custom = properties.customProperties;
        context.load(custom);
        var longstring = "";

        var myTable= "<table><tr><td style='width: 100px; color: red;'>Property</td>";
        myTable+= "<td style='width: 100px; color: red; text-align: right;'>Value</td>";
        myTable+= "<td style='width: 100px; color: red; text-align: right;'>Insert</td></tr>";

          return context.sync().then(function(){
            properties.title = properties.title + " Additional Title Text"; 
            
            for(var i = 0; i < custom.items.length; i++){
              longstring += custom.items[i].key + ": " + custom.items[i].value +"<br>";
              
              myTable+="<tr><td style='width: 100px;'>" + custom.items[i].key + "</td>";
              myTable+="<td style='width: 100px; text-align: right;'>" + custom.items[i].value + "</td>";
              myTable+="<td style='width: 100px; text-align: right;'><input type='radio' name='fieldNameSelection' value='" + custom.items[i].key + "'/></td></tr>";
              
              //myTable+="<td style='width: 100px; text-align: right;'><button id='" + custom.items[i].key + "'>Insert</button></td></tr>";
              
              //$(custom.items[i].key).on("click", {fieldname:custom.items[i].key}, insertfieldxmlname);
            }
            //$('#wordProps').html(longstring);
            myTable+="</table>";
            document.getElementById('wordProps').innerHTML = myTable;
            
         
            
            
        return context.sync();
    });    
});
    }
  
  
  
    function insertfieldxmlname() {

    var fieldname = $('input[name="fieldNameSelection"]:checked').val();
    
    var myOOXMLRequest = new XMLHttpRequest();
    var myXML;
 //   myOOXMLRequest.open('GET', fileName, false);
//    myOOXMLRequest.send();
//    if (myOOXMLRequest.status === 200) {
//        myXML = myOOXMLRequest.responseText;
//    }
    
    
    myXML = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
        <w:body>
          <w:p>
            <w:fldSimple w:instr="DOCPROPERTY ${fieldname} \\* MERGEFORMAT">
              <w:r>
                <w:t>${fieldname}</w:t>
              </w:r>
           </w:fldSimple>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    
    
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });

}
    
      
    
    
    
    function insertfieldxml() {

    var myOOXMLRequest = new XMLHttpRequest();
    var myXML;
 //   myOOXMLRequest.open('GET', fileName, false);
//    myOOXMLRequest.send();
//    if (myOOXMLRequest.status === 200) {
//        myXML = myOOXMLRequest.responseText;
//    }
    
    
    myXML = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
        <w:body>
          <w:p>
            <w:fldSimple w:instr="DOCPROPERTY JOB \\* MERGEFORMAT">
              <w:r>
                <w:t>JOBtxt</w:t>
              </w:r>
           </w:fldSimple>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    
    
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });

}
    
    

    
 
    function insertEmersonQuoteAtSelection() {
        Word.run(function (context) {
 
            // Create a proxy object for the document.
            var thisDocument = context.document;
 
            // Queue a command to get the current selection.
            // Create a proxy range object for the selection.
            var range = thisDocument.getSelection();
 
            // Queue a command to replace the selected text.
            range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);
 
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Ralph Waldo Emerson.');
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }
 
    function insertChekhovQuoteAtTheBeginning() {
        Word.run(function (context) {
 
            // Create a proxy object for the document body.
            var body = context.document.body;
 
            // Queue a command to insert text at the start of the document body.
            body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);
 
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from Anton Chekhov.');
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }
 
    function insertChineseProverbAtTheEnd() {
        Word.run(function (context) {
 
            // Create a proxy object for the document body.
            var body = context.document.body;
 
            // Queue a command to insert text at the end of the document body.
            body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);
 
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added a quote from a Chinese proverb.');
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }
})();
