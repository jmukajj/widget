(function () {
  const template = document.createElement('template');
  template.innerHTML = `
        <style>
        </style>
        
        <div id="root" style="width: 100%; height: 100%;">
          <p><a id="link_href" href="https://www.google.com/" target="_blank">Google</a></p>
          <button id="generateWordBtn">Generate Word Document</button>
        </div>
      `;

  class Main extends HTMLElement {
    constructor() {
      super();
      this._shadowRoot = this.attachShadow({ mode: 'open' });
      this._shadowRoot.appendChild(template.content.cloneNode(true));

      this.generateWordBtn = this._shadowRoot.getElementById('generateWordBtn');
      this.generateWordBtn.addEventListener('click', () => this.generateWordDocument());
    }

    setLink(link) {
      this._link = link;
    }

    setServerSAP(ServerSAP) {
      this._ServerSAP = ServerSAP;
    }

    setODataServiceSAP(ODataService) {
      this._ODataService = ODataService;
    }

    sendPostData(postData) {
      this._postData = postData;
      this.render();
    }

    sendGet() {
      this.render();
    }

    getResponse() {
      return this.Response;
    }

    getLink() {
      return this._link;
    }

    onCustomWidgetResize(width, height) {}

    onCustomWidgetAfterUpdate(changedProps) {}

    onCustomWidgetDestroy() {}

    // Modified render method to fetch data, log it, and generate a Word document
    async render() {
      try {
        console.log("Rendering started...");
        
        // Log postData for debugging purposes
        if (!this._postData) {
          console.error("No post data available!");
          return;
        }
        
        console.log("Post data:", this._postData);

        const url = `https://${this._ServerSAP}/${this._ODataService}`;
        
        // Step 1: Fetch CSRF token
        const getRequest = new XMLHttpRequest();
        getRequest.open('GET', url, true);
        getRequest.setRequestHeader('X-CSRF-Token', 'Fetch');
        getRequest.setRequestHeader('Access-Control-Allow-Methods', 'GET');
        getRequest.setRequestHeader('Access-Control-Allow-Origin', 'https://itsvac-test.eu20.hcs.cloud.sap');
        getRequest.setRequestHeader('Access-Control-Allow-Credentials', true);
        getRequest.setRequestHeader('Access-Control-Expose-Headers', 'X-Csrf-Token,x-csrf-token');
        getRequest.setRequestHeader('Content-Type', 'application/json');
        getRequest.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
        getRequest.withCredentials = true;
        getRequest.send();

        // Step 2: Once the CSRF token is fetched, process the POST request
        getRequest.onreadystatechange = () => {
          if (getRequest.readyState === 4) {
            const csrfToken = getRequest.getResponseHeader('x-csrf-token');
            console.log("CSRF token fetched:", csrfToken);

            // Proceed with POST request
            const data = this._postData; // Data to be posted
            const postRequest = new XMLHttpRequest();
            postRequest.open('POST', url, true);
            postRequest.setRequestHeader('Content-type', 'application/json');
            postRequest.setRequestHeader('Access-Control-Allow-Credentials', true);
            postRequest.setRequestHeader('Cache-Control', 'no-cache');
            postRequest.setRequestHeader("X-Referrer-Hash", window.location.hash);
            postRequest.setRequestHeader('Access-Control-Allow-Origin', 'https://itsvac-test.eu20.hcs.cloud.sap');
            postRequest.setRequestHeader('Access-Control-Allow-Methods', 'POST');
            postRequest.setRequestHeader('X-CSRF-Token', csrfToken);
            postRequest.withCredentials = true;

            postRequest.send(JSON.stringify(data));

            postRequest.onreadystatechange = () => {
              if (postRequest.readyState === 4) {
                if (postRequest.status === 201) {
                  this.Response = JSON.parse(postRequest.responseText);
                  console.log("Post request successful. Response:", this.Response);
                }
              }
            };
          }
        };

        // Step 3: After fetching the data, generate the Word document
        const { CreatedBy, CreatedOn, TotalAmount } = this._postData;
        
        // Use docx library to create Word document
        const doc = new docx.Document({
          sections: [{
            properties: {},
            children: [
              new docx.Paragraph({
                text: "Antrag Details",
                heading: docx.HeadingLevel.HEADING_1,
              }),
              new docx.Paragraph(`Created by: ${CreatedBy}`),
              new docx.Paragraph(`Created on: ${CreatedOn}`),
              new docx.Paragraph(`Total Amount: ${TotalAmount}`),
            ],
          }],
        });

        // Export the document to a blob and trigger download
        docx.Packer.toBlob(doc).then(blob => {
          const link = document.createElement("a");
          link.href = URL.createObjectURL(blob);
          link.download = "AntragDetails.docx";
          link.click();
        });

      } catch (error) {
        console.error("Error during render:", error);
      }
    }

    // New method to generate Word document based on the data
    async generateWordDocument() {
      if (!this._postData) {
        console.error("No post data available");
        return;
      }

      this.render();
    }
  }

  customElements.define('com-sap-sac-jm', Main);
})();
