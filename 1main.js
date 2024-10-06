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
                    <p><a id="link_href" href="#">Update and Save Word Document</a></p>
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
        this._postData = { rows: selectedRowsData };
        console.log("Post Data after population: ", this._postData);
    }

    updateExistingDocument() {
        const data = this._postData;
        if (!data || Object.keys(data).length === 0) {
            alert("No data to update the document");
            return;
        }

        // Fetch the Word template, populate it, and update the document in-memory
        this.fetchLocalTemplate()
            .then(templateBlob => this.populateWordTemplate(templateBlob, data))
            .then((updatedBlob) => {
                alert("Document has been successfully updated!");
                console.log("Document has been successfully updated.");
                // Save the updated document
                saveAs(updatedBlob, 'updated_document.docx');
            })
            .catch(error => {
                console.error("Error updating document:", error);
                alert("Error updating the document. Please check the console for details.");
            });
    }

    fetchLocalTemplate() {
        return new Promise((resolve, reject) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.docx';
            input.style.display = 'none';
            document.body.appendChild(input);

            input.addEventListener('change', (event) => {
                const file = event.target.files[0];
                if (file) {
                    resolve(file);
                } else {
                    reject(new Error("No file selected"));
                }
            });

            input.click();
            document.body.removeChild(input);
        });
    }

    populateWordTemplate(templateBlob, data) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function (event) {
                const arrayBuffer = event.target.result;

                let zip;
                try {
                    zip = new PizZip(arrayBuffer);
                } catch (error) {
                    console.error("Error reading the array buffer with PizZip:", error);
                    return reject(error);
                }

                let doc;
                try {
                    doc = new window.docxtemplater(zip, {
                        paragraphLoop: true,
                        linebreaks: true,
                    });
                } catch (error) {
                    console.error("Error initializing docxtemplater:", error);
                    return reject(error);
                }

                try {
                    console.log("Data passed to docxtemplater:", data);
                    doc.setData(data);
                } catch (error) {
                    console.error("Error setting data for docxtemplater:", error);
                    return reject(error);
                }

                try {
                    doc.render();
                    alert("Document has been successfully populated in memory!");
                    console.log("Document has been successfully populated in memory!");
                } catch (error) {
                    console.error("Error rendering document:", error);
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
