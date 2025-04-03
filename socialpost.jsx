function parseCSV(csvText) {
  // Ensure csvText is a string.
  csvText = csvText.toString();
  var rows = csvText.split("\n");
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    // Convert each row to a string and remove carriage returns and extra spaces
    var rowText = rows[i].toString();
    rowText = rowText.replace(/\r/g, "").replace(/^\s+|\s+$/g, "");
    if (rowText === "") continue; // Skip empty lines
    var cells = rowText.split(",");
    // Clean each cell: remove surrounding quotes and trim spaces
    for (var j = 0; j < cells.length; j++) {
      cells[j] = cells[j].replace(/^"|"$/g, "").replace(/^\s+|\s+$/g, "");
    }
    result.push(cells);
  }
  return result;
}

// ------------------ Main Function ------------------
function insertSpeakerDataInDesign() {
  // Let the user select a CSV file manually.
  var fileRef = File.openDialog("Select CSV file", "*.csv");
  if (!fileRef) {
    alert("No file selected.");
    return;
  }
  
  fileRef.encoding = "UTF-8";
  if (!fileRef.open("r")) {
    alert("Unable to open the file.");
    return;
  }
  
  var csvData = fileRef.read();
  fileRef.close();
  
  var data = parseCSV(csvData);
  if (data.length < 1) {
    alert("No data found in CSV.");
    return;
  }
  
  // Remove header row (if present)
  data.shift();
  
  // Group data by sequence (Column A)
  var sequenceMap = {};
  for (var i = 0; i < data.length; i++) {
    var sequence = data[i][0]; // Column A: Sequence number
    if (!sequenceMap[sequence]) {
      sequenceMap[sequence] = [];
    }
    // Push an array: [Name, Company, Session Title] (Columns B, C, D)
    sequenceMap[sequence].push(data[i].slice(1));
  }
  
  // Ensure a document is open in InDesign
  if (app.documents.length === 0) {
    alert("No document open.");
    return;
  }
  var doc = app.activeDocument;
  var pages = doc.pages;
  
  // Iterate through each sequence group.
  // This mapping assumes: sequence "1" → page index 0, "2" → page index 1, etc.
  for (var seq in sequenceMap) {
    var speakerDataGroup = sequenceMap[seq];
    var pageIndex = parseInt(seq) - 1;
    if (pageIndex < 0 || pageIndex >= pages.length) {
      $.writeln("No page for sequence " + seq);
      continue;
    }
    var page = pages[pageIndex];
    var frames = page.textFrames;
    
    // For multiple speakers on the same sequence, keep track of the current speaker.
    var speakerIndex = 0;
    for (var f = 0; f < frames.length; f++) {
      var frame = frames[f];
      var text = frame.contents.toString();
      // Remove extra spaces from the beginning and end
      text = text.replace(/^\s+|\s+$/g, "");
      
      // Replace the "Session Title" placeholder (once per page)
      if (text === "Session Title" && speakerDataGroup.length > 0) {
        frame.contents = speakerDataGroup[0][2]; // Column D: Session Title
      }
      
      // Replace placeholders for "Name" and "Company"
      if (text.indexOf("Name") !== -1 || text.indexOf("Company") !== -1) {
        if (speakerIndex < speakerDataGroup.length) {
          var speakerName = speakerDataGroup[speakerIndex][0];   // Column B: Name
          var speakerCompany = speakerDataGroup[speakerIndex][1];  // Column C: Company
          var updatedText = text.replace("Name", speakerName).replace("Company", speakerCompany);
          frame.contents = updatedText;
          speakerIndex++;
        } else {
          // If there are more placeholders than speakers, reuse the last speaker's data.
          var last = speakerDataGroup[speakerDataGroup.length - 1];
          var updatedText = text.replace("Name", last[0]).replace("Company", last[1]);
          frame.contents = updatedText;
        }
      }
    }
  }
  
  alert("Data inserted into InDesign.");
}

// ------------------ Run the Main Function ------------------
insertSpeakerDataInDesign();