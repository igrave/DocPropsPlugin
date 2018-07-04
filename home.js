



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
        //        $('#fieldaddname').click(insertfieldxmlname);
                $('#props').click(showProps);
                $('#addprop').click(addProp);
                $('#addpropdirect').click(addPropInsert);
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or greater.');
            }
        });
        showProps(); 
    };
    
    

      
    var gPropNames = [];
    var gPropValues = [];
    
    
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
    //    myTable+= "<td style='width: 100px; color: red; text-align: right;'>Button</td></tr>";

          return context.sync().then(function(){
            properties.title = properties.title + " Additional Title Text"; 
            
            
            gPropValues.length = 0;
            gPropNames.length = 0;
            
            for(var i = 0; i < custom.items.length; i++){
              longstring += custom.items[i].key + ": " + custom.items[i].value +"<br>";
              
              myTable+="<tr><td style='width: 100px;'>" + custom.items[i].key + "</td>";
              myTable+="<td style='width: 100px; text-align: right;'>" + custom.items[i].value + "</td>";
           //   myTable+="<td style='width: 100px; text-align: right;'><input type='radio' name='fieldNameSelection' value='" + custom.items[i].key + "'/></td></tr>";
              
              myTable+="<td style='width: 100px; text-align: right;'><button id='" + custom.items[i].key + "'>Insert</button></td></tr>";
              
              gPropValues[i] = custom.items[i].value;
              gPropNames[i] = custom.items[i].key;
              
            }
            //$('#wordProps').html(longstring);
            myTable+="</table>";
            document.getElementById('wordProps').innerHTML = myTable;
            
        for(var j = 0; j < gPropNames.length; j++){
          $('#'+custom.items[j].key).on("click",
              {name: custom.items[j].key, value: custom.items[j].value},
              insertfieldxmlpars);
        } 

        return context.sync();
    });    
});
    }
  
  
  
  
  
  function deleteProp(name){
    Word.run(function(context){
        var properties = context.document.properties;
        context.load(properties);
        var custom = properties.customProperties;
        context.load(custom);
        
        var item = custom.getItem(name);
        item.delete();
        
        return context.sync().then(showProps);
        
        
    });
  }
  
  
  
  function addProp(){
     Word.run(function(context){
        var properties = context.document.properties;
        context.load(properties);
        var custom = properties.customProperties;
        context.load(custom);
        
        var name = $('input[name="newName"]').val();
        var value = $('input[name="newValue"]').val();
        
        custom.add(name, value);
        
        $('input[name="newName"]').val("");
        $('input[name="newValue"]').val("");
        
        return context.sync().then(showProps);
        
        
    });
  }
    
    
    
  function addPropInsert(){
      console.log("Trying");
      addProp();
      console.log("addProp is finished");
      insertfield($('input[name="newName"]').val(), $('input[name="newValue"]').val());
      console.log("insert is finished");
    }
    
  
  
  function insertfieldxmlpars(event) {

   

  var fieldname = event.data.name;
  var fieldvalue = event.data.value;
  insertfield(fieldname, fieldvalue);
}






  
    function insertfield2013(fieldname, fieldvalue) {
  var myXML;

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
                <w:t>${fieldvalue}</w:t>
              </w:r>
           </w:fldSimple>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    
    
    Office.context.document.setSelectedDataAsync(
      myXML,
      { coercionType: 'ooxml' },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Action failed with error: " + asyncResult.error.message);
          }
        });
        
        
    Word.run(function (context) {
        var range = context.document.getSelection();
        range.select('end');
        return context.sync();
    });

}
    










  function insertfield2016(fieldname, fieldvalue) {
  var myXML;

   myXML = "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n  <pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\" pkg:padding=\"512\">\n    <pkg:xmlData>\n      <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n        <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n      </Relationships>\n    </pkg:xmlData>\n  </pkg:part>\n  <pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\">\n    <pkg:xmlData>\n      <w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" >\n        <w:body>\n          <w:p>\n            <w:fldSimple w:instr=\"DOCPROPERTY " + fieldname + " \\* MERGEFORMAT\">\n              <w:r>\n                <w:t>" + fieldvalue + "</w:t>\n              </w:r>\n           </w:fldSimple>\n          </w:p>\n        </w:body>\n      </w:document>\n    </pkg:xmlData>\n  </pkg:part>\n</pkg:package>";
    
    
      Word.run(function (context) {
 
           var thisDocument = context.document;
           var range = thisDocument.getSelection();
 
            // Queue a command to replace the selected text.
            range.insertOoxml(myXML, Word.InsertLocation.before);
 
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Added XML for DOCPROPERTY '+ fieldname);
            });
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }


    
  
  
    function insertfield2013(fieldname, fieldvalue) {
  var myXML;

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
                <w:t>${fieldvalue}</w:t>
              </w:r>
           </w:fldSimple>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    
    
    Office.context.document.setSelectedDataAsync(
      myXML,
      { coercionType: 'ooxml' },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log("Action failed with error: " + asyncResult.error.message);
          }
        });

}
    
  
  
  

 
 
})();
