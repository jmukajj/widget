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
        this._postData = [];

        this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
            this.updateExistingDocument();
        });

        // Load the libraries in the correct order
        this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
            .then(() => {
                console.log("FileSaver.js library loaded successfully!");
                return this.loadScript('https://cdn.jsdelivr.net/npm/pizzip@3.1.1/dist/pizzip.min.js');
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

    sendPostData(selectedRowsData) {
        console.log("Received selected rows data: ", selectedRowsData);
        if (!selectedRowsData || selectedRowsData.length === 0) {
            console.error("No data provided in selected rows", selectedRowsData);
            alert("No data to update the document");
            return;
        }
        this._postData = selectedRowsData;
        console.log("Post Data after population: ", this._postData);
    }

    updateExistingDocument() {
        const data = this._postData;
        if (!data || data.length === 0) {
            alert("No data to update the document");
            return;
        }

        // Fetch the Word template, populate it, and allow the user to download the updated document
        this.fetchTemplateFromURL(this.templateURL)
            .then(templateBlob => this.populateWordTemplate(templateBlob, data))
            .then((updatedBlob) => {
                // Trigger the download of the updated document
                saveAs(updatedBlob, 'updated_document.docx');
                alert("Document has been successfully updated and downloaded!");
                console.log("Document has been successfully updated and downloaded.");
            })
            .catch(error => {
                console.error("Error updating document:", error);
                alert("Error updating the document. Please check the console for details.");
            });
    }

    fetchTemplateFromURL(url) {
        return fetch(url)
            .then(response => {
                if (!response.ok) {
                    throw new Error('Network response was not ok ' + response.statusText);
                }
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

                // Map data to Word template
                const sanitizedData = {
                    rows: data.map((row, index) => {
                        // Extract checkbox values from the row data (assuming they come as an array)
                        const [ID_1, ID_2, ID_3, ID_4, ID_5, ID_6, ID_7, ID_8] = row.checkboxValues || [];

                        return {
                            AccountDescription: row.AccountDescription || '',
                            Antrag: row.Antrag || '',
                            ID_1: ID_1 || '☐',  // Default to unchecked if not provided
                            ID_2: ID_2 || '☐',
                            ID_3: ID_3 || '☐',
                            ID_4: ID_4 || '☐',
                            ID_5: ID_5 || '☐',
                            ID_6: ID_6 || '☐',
                            ID_7: ID_7 || '☐',
                            ID_8: ID_8 || '☐'
                        };
                    })
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
