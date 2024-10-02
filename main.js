<script type="text/javascript">
  async function runPythonAndModifyDocument(data, documentPath, outputPath, oldChar, newChar) {
    // Load Pyodide
    const pyodide = await loadPyodide();
    
    // Define the Python code
    const pythonCode = `
      from docx import Document

      def replace_checkboxes(data, document_path, output_path, old_char, new_char):
          # Open the existing document
          doc = Document(document_path)
          for id in data:
              for paragraph in doc.paragraphs:
                  if id in paragraph.text:
                      for index, run in enumerate(paragraph.runs):
                          if id in run.text:
                              paragraph.runs[index-2].text = paragraph.runs[index-2].text.replace(old_char, new_char)
          doc.save(output_path)

      # Run the function with the passed arguments
      replace_checkboxes(${data}, "${documentPath}", "${outputPath}", "${oldChar}", "${newChar}")
    `;

    // Run the Python code
    try {
      await pyodide.runPython(pythonCode);
      console.log("Python script executed successfully!");
    } catch (error) {
      console.error("Error executing Python script:", error);
    }
  }

  class Main extends HTMLElement {
    constructor() {
      super();
      console.log('Widget initialized');
      this._shadowRoot = this.attachShadow({ mode: 'open' });
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

      this._shadowRoot.appendChild(template.content.cloneNode(true));
      this._postData = {};

      this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
        this.generateAndDownloadDocument();
      });

      // Load docx library and Pyodide
      this.loadScript('https://cdn.jsdelivr.net/pyodide/v0.18.1/full/pyodide.js')
        .then(() => {
          console.log("Pyodide library loaded successfully!");
        })
        .catch((error) => {
          console.error("Error loading the Pyodide library:", error);
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

    async generateAndDownloadDocument() {
      const data = this._postData;
      if (!data || Object.keys(data).length === 0) {
        alert("No data to generate document");
        return;
      }

      // Define the required parameters for the Python script
      const documentPath = "https://github.com/jmukajj/widget/raw/refs/heads/main/template.docx'"; // Adjust the path as needed
      const outputPath = "https://github.com/jmukajj/widget/raw/refs/heads/main/template.docx'";
      const oldChar = "\\uf06f";
      const newChar = "\\uf0fd";
      const dataToReplace = Object.values(data); // Assuming this._postData contains values like ["{{ID:1}}", "{{ID:2}}"]

      // Run Python script to modify the Word document using Pyodide
      await runPythonAndModifyDocument(dataToReplace, documentPath, outputPath, oldChar, newChar);
    }
  }

  customElements.define('com-sap-sac-jm', Main);
</script>
