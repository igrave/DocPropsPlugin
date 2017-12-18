
//https://stackoverflow.com/questions/39826654/how-do-we-add-field-code-for-word-using-word-javascript-api

// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    var range = context.document.getSelection();

    // Queue a commmand to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'>
         <pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'>
             <pkg:xmlData>
                 <Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>
             </pkg:xmlData>
         </pkg:part>
         <pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'>
             <pkg:xmlData>
                 <w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' >
                 <w:body>
                 <w:p w:rsidR="00000000" w:rsidRDefault="0043114D">
                     <w:r>
                        <w:fldChar w:fldCharType="begin"/>
                     </w:r>
                     <w:r>
                        <w:instrText xml:space="preserve"> PAGE  \* Arabic  \* MERGEFORMAT </w:instrText>
                     </w:r>
                     <w:r>
                        <w:fldChar w:fldCharType="separate"/>
                     </w:r>
                     <w:r>
                         <w:rPr>
                            <w:noProof/>
                         </w:rPr>
                         <w:t>1</w:t>
                     </w:r>
                     <w:r>
                         <w:fldChar w:fldCharType="end"/>
                     </w:r>
                     </w:p>
                 </w:body>
                 </w:document>
             </pkg:xmlData>
         </pkg:part>
    </pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the range.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
