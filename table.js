var selektion = Table_1.getSelections(); // Get the selected row data
console.log(selektion); // Log the entire selection

if (selektion && selektion.length > 0) {
    if (selektion[0] && selektion[0]["P2R_G_ANTRAG"] !== undefined) {
        // Retrieve the name of the selected Account
        var accountValue = selektion[0]["P2R_G_ANTRAG"];  // Extract the Account value from the selected row
        req_desc = Table_1.getDataSource().getMember("P2R_G_ANTRAG", accountValue).description;  // Fetch the description using the extracted Account value
        console.log(req_desc);
            
        // Retrieve the Antrag status
        var antragStatusValue = selektion[0]["P2R_G_ANTRAG.Antrag_Status"]; // Extract the Antrag status value from the selected row
        antragStatus = Table_1.getDataSource().getMember("P2R_G_ANTRAG.Antrag_Status", antragStatusValue).description;
        console.log(antragStatus);
        
        // Retrieve the Antrag description
        var antragDesc = selektion[0]["P2R_G_ANTRAG.Antragsbeschreibung"]; // Extract the description value
        console.log(antragDesc);	
		
		// besch_bereich logic to retrieve the description based on value
		var besch_bereich_id = selektion[0]["P2R_G_BESCHAFFENDER_BEREICH"]; // Extract the Beschaffender_Bereich ID
		console.log(besch_bereich_id);

		if (besch_bereich_id) {
			var besch_bereich_description = Table_1.getDataSource().getMember("P2R_G_BESCHAFFENDER_BEREICH", besch_bereich_id).description; // Fetch the description
			console.log(besch_bereich_description);

			// Apply your logic based on the description
			if (besch_bereich_description === "BFMB") {
				besch_bereich_id = "ID_4";
			} else if (besch_bereich_description === "IT") {
				besch_bereich_id = "ID_5";
			} else if (besch_bereich_description === "PE") {
				besch_bereich_id = "ID_6";
			} else if (besch_bereich_description === "GFPPH") {
				besch_bereich_id = "ID_13";
			} else {
				besch_bereich_id = "Unassigned";
			}
		} else {
			besch_bereich_id = "Unassigned";
		}

		console.log(besch_bereich_id);
		
        // Retrieve raw account values from the result set and apply condition for ID assignment
        var arr = Table_1.getDataSource().getResultSet();
        var arrvalue = ArrayUtils.create(Type.number);
        var accountID = "Unassigned"; // Default if no conditions match
        
        for (var i = 0; i < arr.length; i++) {
            var rawValue = arr[i]['Account'].rawValue;
            var numberValue = ConvertUtils.stringToNumber(rawValue);
            arrvalue.push(numberValue);
        }
        console.log(arrvalue);  // Log the array values
		
        // Determine the ID based on the account value
        if (arrvalue[0] < 54540) {
            accountID = "ID_7"; // < € 54.540,00
        } else if (arrvalue[0] >= 54540.01 && arrvalue[0] <= 109080) {
            accountID = "ID_8"; // € 54.540,01 bis € 109.080,00
        } else if (arrvalue[0] >= 109080.01 && arrvalue[0] <= 214500) {
            accountID = "ID_9"; // € 109.080,01 bis € 214.500,00
        } else if (arrvalue[0] > 214500) {
            accountID = "ID_10"; // Verwaltungsrat gem. § 41 Abs. 1 Z2 BO (in sonstigen Fällen)
        } else {
            accountID = "ID_7"; // Büro gem. § 41 Abs. 1 Z1 BO in folgenden Fällen
        }

        // Log and create the object with the relevant data
        console.log(accountID);
        var postData = {
            Antrag: req_desc, // Include the Account name
            AntragStatus: antragStatus, // Include the Antrag status
            AntragDescription: antragDesc, // Include the Antrag description
            Besch_Berich_desc: besch_bereich_description, // Include the selected Besch_Bereich Desc
            AccountID: accountID // Include the selected Account ID based on value condition
        };
		
        // Send the extracted data to the widget
        Widget_1.sendPostData(postData); // Send the data to the widget
    } else {
        console.log("No valid data found in the selected row. Missing 'P2R_G_ANTRAG' key.");
    }
} else {
    console.log("No rows were selected or the selection returned undefined.");
}
