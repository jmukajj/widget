(function () {
  const template = document.createElement('template');
  template.innerHTML = `
       <style>
        #root {
            width: 300px;
            justify-content: flex-start;
            align-items: flex-start;
            height: 100vh;
        }
        .link-container {
            position: relative;
            display: flex;
            flex-direction: column;
            justify-content: left;
            left: 7px;
            align-items: left;
            padding: 20px;
            border: 0.5px solid black;
            background-color: #FCFCFC;
            box-shadow: 0 4px 8px rgba(0, 0, 0, .3);
        }
        .link-container::before {
            content: '';
            position: absolute;
            left: -5.5px;
            top: 50%;
            transform: translateY(-50%) rotate(135deg);
            border: solid black;
            border-width: 0 0.5px 0.5px 0;
            display: inline-block;
            padding: 5px;
            background-color: #FCFCFC;
        }
        .link {
            text-decoration: none;
            color: #5E97C4;
            font-family: Arial, sans-serif;
            margin-bottom: 10px;
            display: block;
        }
       </style>

       <div id="root">
          <div class="link-container" id="links-container">
            <p><a id="link_href" href="#" target="_blank">Download Word Document</a></p>
          </div>
       </div>
  `;

  class Main extends HTMLElement {
    constructor () {
      super();
      console.log('Widget initialized');
      this._shadowRoot = this.attachShadow({ mode: 'open' });
      this._shadowRoot.appendChild(template.content.cloneNode(true));
      this.Response = null;
      this._postData = {};

      this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
        this.generateAndUploadDocument();
      });

      // Load the libraries in the correct order
      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
        .then(() => {
          console.log("FileSaver.js library loaded successfully!");
          return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/pizzip/3.0.6/pizzip.min.js');
        })
        .then(() => {
          console.log("PizZip library loaded successfully!");
          return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.21.2/docxtemplater.min.js');
        })
        .then(() => {
          console.log("docxtemplater library loaded successfully!");
        })
        .catch((error) => {
          console.error("Error loading a library:", error);
        });
    }

    loadScript(url) {
      return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = url;
        script.async = false;
        script.onload = () => resolve();
        script.onerror = () => reject(`Failed to load script: ${url}`);
        document.head.appendChild(script);
      });
    }

    sendPostData(selectedRowData) {
      console.log("Received selected row data: ", selectedRowData);
      if (!selectedRowData || Object.keys(selectedRowData).length === 0) {
        console.error("No data provided in selected row", selectedRowData);
        alert("No data to generate document");
        return;
      }
      this._postData = selectedRowData;
      console.log("Post Data after population: ", this._postData);
    }

    generateAndUploadDocument() {
      const data = this._postData;
      if (!data || Object.keys(data).length === 0) {
        alert("No data to generate document");
        return;
      }

      fetchWordTemplate()
        .then(templateBlob => populateWordTemplate(templateBlob, data))
        .then(populatedDocument => uploadToGithub(populatedDocument, 'populated_document.docx'))
        .catch(error => console.error("Error generating and uploading document:", error));
    }
  }

  customElements.define('com-sap-sac-jm', Main);

  // Fetch the Word Template from your GitHub Repo
  async function fetchWordTemplate() {
    const response = await fetch('https://jmukajj.github.io/widget/template.docx');
    if (!response.ok) {
      throw new Error('Failed to fetch the Word template');
    }
    return await response.blob();
  }

  // Populate the Word Template
  function populateWordTemplate(templateBlob, data) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = function (event) {
        const arrayBuffer = event.target.result;
        const zip = new PizZip(arrayBuffer);

        let doc;
        try {
          doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
          });
        } catch (error) {
          return reject(error);
        }

        // Set the template variables
        doc.setData({
          AccountDescription: data.AccountDescription || '',
          Antrag: data.Antrag || '',
        });

        try {
          doc.render();
        } catch (error) {
          return reject(error);
        }

        const out = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });
        resolve(out);
      };
      reader.readAsArrayBuffer(templateBlob);
    });
  }

  // Upload the Document to GitHub
  async function uploadToGithub(fileBlob, fileName) {
    const fileReader = new FileReader();
    fileReader.onloadend = async function () {
      const content = fileReader.result.split(',')[1];
      const data = {
        message: `Upload ${fileName}`,
        content: content,
        branch: "main"
      };

      // GitHub API Request
      const response = await fetch(`https://jmukajj.github.io/widget/template.docx`, {
        method: "PUT",
        headers: {
          "Authorization": `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`, // Replace with your GitHub Token
          "Content-Type": "application/json"
        },
        body: JSON.stringify(data)
      });

      if (response.ok) {
        console.log(`${fileName} uploaded successfully.`);
      } else {
        console.error('Upload failed:', response.statusText);
      }
    };
    fileReader.readAsDataURL(fileBlob);
  }
})();
