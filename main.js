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
                height: auto;
            }
            .link-container {
                position: relative;
                display: flex;
                flex-direction: column;
                justify-content: flex-start;
                align-items: flex-start;
                padding: 20px;
                border: 0.5px solid black;
                background-color: #FCFCFC;
                box-shadow: 0 4px 8px rgba(0, 0, 0, .3);
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
                    <p><a id="link_href" href="#">Download Updated Word Document</a></p>
                </div>
            </div>
        `;
        
        this._shadowRoot.appendChild(template.content.cloneNode(true));
        this._postData = null; // Now expect a single object instead of an array

        // Set the correct template URL from GitHub
        this.templateURL = "https://jmukajj.github.io/widget/template.docx"; //  GitHub URL
        
        this._shadowRoot.getElementById('link_href').addEventListener('click', (event) => {
            event.preventDefault();  // Prevent default link behavior
            console.log("Download link clicked, attempting to update the document...");
            this.updateExistingDocument();
        });

        // Load external libraries in sequence
        this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
            .then(() => {
                console.log("FileSaver.js loaded");
                return this.loadScript('https://cdn.jsdelivr.net/npm/pizzip@3.1.1/dist/pizzip.min.js');
            })
            .then(() => {
                console.log("PizZip loaded");
                return this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.21.2/docxtemplater.min.js');
            })
            .then(() => console.log("docxtemplater loaded"))
            .catch(error => console.error("Error loading a library:", error));
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

    // Updated to receive a single object instead of an array
    sendPostData(selectedRowData) {
        console.log("Received data in widget:", selectedRowData);
        if (!selectedRowData || Object.keys(selectedRowData).length === 0) {
            alert("No data to update the document");
            return;
        }
        this._postData = selectedRowData; // Store the single object in _postData
        console.log("Post data set in widget:", this._postData);
    }

    // Now process the single object
    updateExistingDocument() {
        const data = this._postData; // This is the object you passed via sendPostData
        console.log("Updating document with data:", data);
        if (!data || Object.keys(data).length === 0) {
            alert("No data to update the document");
            return;
        }

        this.fetchTemplateFromURL(this.templateURL)
            .then(templateBlob => this.populateWordTemplate(templateBlob, data))
            .then(updatedBlob => {
                console.log("Document updated, initiating download");
                saveAs(updatedBlob, 'updated_document.docx');
            })
            .catch(error => {
                console.error("Error updating document:", error);
                alert("An error occurred while updating the document. Check the console for details.");
            });
    }

    fetchTemplateFromURL(url) {
        console.log("Fetching template from URL:", url);
        return fetch(url).then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok ' + response.statusText);
            }
            console.log("Template fetched successfully");
            return response.blob();
        });
    }

    populateWordTemplate(templateBlob, data) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = event => {
                const arrayBuffer = event.target.result;
                let zip;

                try {
                    zip = new PizZip(arrayBuffer);
                } catch (error) {
                    return reject(error);
                }

                let doc;
                try {
                    doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });
                } catch (error) {
                    return reject(error);
                }

                // Map the single object data to the Word template placeholders
                const sanitizedData = {
                    AccountDescription: data.Antrag || '',
                    AntragStatus: data.AntragStatus || '',
                    AntragDescription: data.AntragDescription || '',
                    ID_4: (data.Besch_Berich_desc === "BFMB") ? '☑' : '☐',
                    ID_5: (data.Besch_Berich_desc === "IT") ? '☑' : '☐',
                    ID_6: (data.Besch_Berich_desc === "PE") ? '☑' : '☐',
                    ID_7: (data.AccountID === "ID7") ? '☑' : '☐',
                    ID_8: (data.AccountID === "ID8") ? '☑' : '☐',
                    ID_9: (data.AccountID === "ID9") ? '☑' : '☐',
                    ID_10: (data.AccountID === "ID10") ? '☑' : '☐'
                };

                try {
                    doc.setData(sanitizedData);
                    doc.render();
                } catch (error) {
                    return reject(error);
                }

                const updatedBlob = doc.getZip().generate({
                    type: 'blob',
                    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                });

                resolve(updatedBlob);
            };

            reader.readAsArrayBuffer(templateBlob);
        });
    }
}

customElements.define('com-sap-sac-jm', Main);
