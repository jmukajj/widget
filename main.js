(function () {
  const template = document.createElement('template');
  template.innerHTML = `
    <style>
    </style>
    <div id="root" style="width: 100%; height: 100%;">
      <p><a id="link_href" href="https://www.google.com/" target="_blank">Google</a></p>
    </div>
  `;

  class Main extends HTMLElement {
    constructor () {
      super();
      this._shadowRoot = this.attachShadow({ mode: 'open' });
      this._shadowRoot.appendChild(template.content.cloneNode(true));
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

    async render() {
      const url = `https://${this._ServerSAP}/${this._ODataService}`;

      // Fetch CSRF token first (GET request)
      var xhrGet = new XMLHttpRequest();
      xhrGet.open('GET', url, true);
      xhrGet.setRequestHeader('X-CSRF-Token', 'Fetch');
      xhrGet.setRequestHeader('Content-Type', 'application/json');
      xhrGet.withCredentials = true;
      xhrGet.send();

      xhrGet.onreadystatechange = () => {
        if (xhrGet.readyState === 4 && xhrGet.status === 200) {
          const __XCsrfToken = xhrGet.getResponseHeader('x-csrf-token');

          if (this._postData) {
            const data = this._postData; // Data to be posted

            // Step 2. Send POST request with the retrieved CSRF token
            var xhr = new XMLHttpRequest();
            xhr.open('POST', url, true);
            xhr.setRequestHeader('Content-type', 'application/json');
            xhr.setRequestHeader('X-CSRF-Token', __XCsrfToken);
            xhr.withCredentials = true;
            xhr.send(JSON.stringify(data));

            xhr.onreadystatechange = () => {
              if (xhr.readyState == 4 && xhr.status == 201) {
                this.Response = JSON.parse(xhr.responseText);
                // Logic to generate Word document using the response data
                this.generateWordDoc(this.Response);
              }
            };
          }
        }
      };
    }

    generateWordDoc(data) {
      const docContent = `
        Antrag Details:
        Created by: ${data.CreatedBy}
        Created on: ${data.CreatedOn}
        Total Amount: ${data.TotalAmount}
      `;

      // Use a library like jsPDF, docx, or similar to generate the Word document
      const doc = new docx.Document({
        sections: [
          {
            properties: {},
            children: [
              new docx.Paragraph(docContent)
            ]
          }
        ]
      });

      docx.Packer.toBlob(doc).then((blob) => {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "Antrag_Details.docx";
        link.click();
      });
    }
  }

  customElements.define('com-sap-sac-jm', Main);
})();
