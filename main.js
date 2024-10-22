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

        // Set the correct template URL from GitHub
        this.templateURL = "https://jmukajj.github.io/widget/template.docx"; //  GitHub URL
        
        this._shadowRoot.getElementById('link_href').addEventListener('click', () => {
            event.preventDefault();
            console.log("Download link clicked, attempting to update the document...");
            this.updateExistingDocument();
        });

        // Load external libraries in sequence
        this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js')
            .then(() => {
                console.log("FileSaver.js loaded");
                return this.loadScript('https://cdn.jsdelivr.net/npm/pizzip@3.1.1/dist/pizzip.min.js');
            })
            .catch(error => console.error("Error loading a library:", error));
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

    sendPostData(selectedRowsData) {
        console.log("Received data in widget:", selectedRowsData);
        if (!selectedRowsData || selectedRowsData.length === 0) {
            alert("No data to update the document");
            return;
        }
        this._postData = selectedRowsData;
        console.log("Post data set in widget:", this._postData);
    }

    updateExistingDocument() {
        const data = this._postData;
        console.log("Updating document with data:", data);
        if (!data || data.length === 0) {
            alert("No data to update the document");
            return;
        }

        this.fetchTemplateFromURL(this.templateURL)
            .then(templateBlob => this.populateWordTemplate(templateBlob, data))
            .then(updatedBlob => {
                console.log("Document updated, initiating download");
                saveAs(updatedBlob, 'updated_document.docx');
            })
            .catch(error => console.error("Error updating document:", error));
            alert("An error occurred while updating the document. Check the console for details.");
            });
        } catch (error) {
            console.error("Unexpected error in updateExistingDocument:", error);
            alert("Unexpected error. Check the console for more details.");
        }
    }

    this.fetchTemplateFromURL(this.templateURL)
        .then(templateBlob => {
            if (!templateBlob) {
                throw new Error("Failed to fetch the template. Please check the template URL.");
            }
            return this.populateWordTemplate(templateBlob, data);
        })
        .catch(error => {
            console.error("Error fetching or processing the template:", error);
            alert("There was an issue fetching or processing the template.");
        });

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

                const sanitizedData = {
                    rows: data.map(row => {
                        const checkboxTrue = '☑';  // Checked checkbox symbol (Unicode U+2611)
                        const checkboxFalse = '☐'; // Unchecked checkbox symbol (Unicode U+2610)

                        let ID_4, ID_5, ID_6, ID_7, ID_8, ID_9, ID_10;

                        if (row.Besch_Berich_ID === "ID4") {
                            ID_4 = checkboxTrue;
                        } else if (row.Besch_Berich_ID === "ID5") {
                            ID_5 = checkboxTrue;
                        } else if (row.Besch_Berich_ID === "ID6") {
                            ID_6 = checkboxTrue;
                        }

                        if (row.AccountID === "ID7") {
                            ID_7 = checkboxTrue;
                        } else if (row.AccountID === "ID8") {
                            ID_8 = checkboxTrue;
                        } else if (row.AccountID === "ID9") {
                            ID_9 = checkboxTrue;
                        } else if (row.AccountID === "ID10") {
                            ID_10 = checkboxTrue;
                        }

                        return {
                            AccountDescription: row.Antrag || '',
                            AntragStatus: row.AntragStatus || '',
                            AntragDescription: row.AntragDescription || '',
                            ID_4: ID_4 || checkboxFalse,
                            ID_5: ID_5 || checkboxFalse,
                            ID_6: ID_6 || checkboxFalse,
                            ID_7: ID_7 || checkboxFalse,
                            ID_8: ID_8 || checkboxFalse,
                            ID_9: ID_9 || checkboxFalse,
                            ID_10: ID_10 || checkboxFalse
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
