(function () {
  const template = document.createElement('template');
  template.innerHTML = `
        <style>
        </style>
        
        <div id="root" style="width: 100%; height: 100%;">
          <button id="sendAntragDataBtn">Send Antrag Data</button>
        </div>
      `

  class Main extends HTMLElement {
    constructor () {
      super();
      this._shadowRoot = this.attachShadow({ mode: 'open' });
      this._shadowRoot.appendChild(template.content.cloneNode(true));

      // Object to store the selected Antrag data
      this.selectedAntrag = {
        createdBy: "",
        createdOn: "",
        totalAmount: 0
      };

      // Variables to store link, Server SAP, and OData Service
      this._link = "";
      this._serverSAP = "";
      this._ODataService = "";

      // Add event listener for the button to send Antrag data
      this._shadowRoot.getElementById('sendAntragDataBtn').addEventListener('click', () => {
        this.fetchAntragData();  // Fetch selected Antrag data
      });
    }

    // Fetch selected Antrag data from the table
    fetchAntragData() {
      // Assuming 'Table_1' contains the Antrag data and the user has selected a row
      var selectedData = Table_1.getSelections(); // Retrieve the selected Antrag from the table

      if (selectedData.length > 0) {
        var antragData = {
          createdBy: selectedData[0].createdBy,  // Replace with actual field name for 'Created by'
          createdOn: selectedData[0].createdOn,  // Replace with actual field name for 'Created on'
          totalAmount: selectedData[0].totalAmount  // Replace with actual field name for 'Total Amount'
        };

        // Set the Antrag data in the widget
        this.setAntragData(antragData);
        this.sendPostData(antragData);  // Send the data
      } else {
        console.log("No Antrag selected");
      }
    }

    // Set the Antrag data
    setAntragData(antragData) {
      this.selectedAntrag = antragData;
      console.log("Antrag data set:", antragData);
    }

    // Set the Link value (from setLink in widget.json)
    setLink(link) {
      this._link = link;
      console.log("Link set:", link);
    }

    // Get the Link value (from getLink in widget.json)
    getLink() {
      return this._link;
    }

    // Set the Server SAP value (from setServerSAP in widget.json)
    setServerSAP(serverSAP) {
      this._serverSAP = serverSAP;
      console.log("Server SAP set:", serverSAP);
    }

    // Set the OData Service SAP value (from setODataServiceSAP in widget.json)
    setODataServiceSAP(ODataService) {
      this._ODataService = ODataService;
      console.log("OData Service SAP set:", ODataService);
    }

    // Send post data
    sendPostData(postData) {
      this._postData = postData;  // Store the post data in the widget instance
      console.log("Post Data to be sent:", postData);

      this.render();  // Call render to simulate sending the data
    }

    // Render or perform the actual sending of data
    async render() {
      console.log("Data to be posted:", this._postData);

      // Here, you can add logic to send the data to an external service via HTTP
      const request = new XMLHttpRequest();
      const url = "https://your-api-endpoint.com";  // Replace with your real API endpoint
      request.open("POST", url, true);
      request.setRequestHeader("Content-Type", "application/json");
      request.onreadystatechange = function () {
        if (request.readyState === 4 && request.status === 200) {
          console.log("Data posted successfully");
        }
      };
      request.send(JSON.stringify(this._postData));  // Send the Antrag data as JSON
    }

    onCustomWidgetAfterUpdate(changedProps) {
      // Handle updates to the custom widget
    }

    onCustomWidgetDestroy() {
      // Clean up when the custom widget is destroyed
    }
  }

  customElements.define('com-sap-sac-jm', Main);
})();
