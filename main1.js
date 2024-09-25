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

      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docx/7.1.0/docx.umd.min.js')
        .then(() => {
          if (window.docx) {
            console.log("docx library loaded successfully!", window.docx);
          } else {
            console.error("docx library failed to load.");
          }
        })
        .catch((error) => {
          console.error("Error loading docx library:", error);
        });

      this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
        .then(() => {
          console.log("FileSaver.js library loaded successfully!");
        })
        .catch((error) => {
          console.error("Error loading FileSaver.js library:", error);
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
    const templateBlob = await response.blob();
    return templateBlob;
  }

  // Populate the Word Template
  function populateWordTemplate(templateBlob, data) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = function (event) {
        const arrayBuffer = event.target.result;
        const doc = new window.docx.Document(arrayBuffer);

        // Loop through paragraphs and replace placeholders
        doc.getSections()[0].getChildren().forEach((paragraph) => {
          const text = paragraph.getText();
          if (text.includes('{{AccountDescription}}')) {
            paragraph.replaceText('{{AccountDescription}}', data.AccountDescription);
          }
          if (text.includes('{{Antrag}}')) {
            paragraph.replaceText('{{Antrag}}', data.Antrag);
          }
        });

        // Generate the populated document
        window.docx.Packer.toBlob(doc).then(blob => {
          resolve(blob);
        }).catch(error => {
          reject(error);
        });
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
      const response = await fetch(`https://api.github.com/repos/jmukajj/widget/contents/${fileName}`, {
        method: "PUT",
        headers: {
          "Authorization": `token YOUR_GITHUB_TOKEN`, // Replace with your GitHub Token
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
