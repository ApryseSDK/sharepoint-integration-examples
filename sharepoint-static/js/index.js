console.log('instantiating webviewer');
let i;

function validateQueryParam(urlParams) {
    const necessaryParams = ['uniqueId', 'tempAuth', 'filename'];
        let result = true;
        necessaryParams.forEach(paramKey => {
        if (!urlParams.get(paramKey)) {
            result = false;
        }
        });
        return result;
    }


WebViewer({
    path: 'js/lib', // path to the PDFTron 'lib' folder on your server
    uiPath: './ui/index.aspx' // make sure to indicate index.aspx instead of index.html, otherwise the UI won't be loaded
  }, document.getElementById('viewer'))
  .then(instance => {
    i = instance;
    createSaveFileButton(instance);
    createSavedModal(instance);
    const urlParams = new URLSearchParams(window.location.search);
    const validateQueryParamResult = this.validateQueryParam(urlParams);
    if (validateQueryParamResult) {
      const uniqueId = urlParams.get("uniqueId");
      const tempAuth = urlParams.get("tempAuth");
      const filename = urlParams.get("filename");
      const newPathnameArray = window.location.pathname.split('/').slice(0, 3);
      const newPathname = newPathnameArray.join('/');
      const domain = window.location.origin;
      const domainUrl = `${domain}${newPathname}`;
      const docUrl = `${domainUrl}/_layouts/15/download.aspx?UniqueId=${uniqueId}&Translate=false&tempauth=${tempAuth}&ApiVersion=2.0`;
      instance.UI.loadDocument(docUrl, {filename});
    } else {
      alert('Please open the webviewer with proper document queries.')
    }
  });

function createSaveFileButton(instance) {
  instance.setHeaderItems(function(header) {
    const saveFileButton = {
      type: 'actionButton',
      dataElement: 'saveFileButton',
      title: 'Save file to sharepoint',
      img: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 0h24v24H0z" fill="none"/><path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/></svg>',
      onClick: async function() {
        const searchParams = new URLSearchParams(window.location.search);
        const folderName = searchParams.get('foldername');
        const fileName = searchParams.get('filename');
        instance.openElement('loadingModal');
        await saveFile(instance, folderName, fileName);
        instance.closeElements(['loadingModal']);
        instance.openElement('savedModal');
      }
    }
    header.get('viewControlsButton').insertBefore(saveFileButton);
  });
}

async function getFormDigestValue() {
  try {
    const resp = await fetch(`${window.location.origin}/sites/5s4vrg/_api/contextinfo`, {
      method: 'POST',
      headers: {
        'Accept': 'application/json; odata=verbose'
      },
    });
    const respJson = await resp.json();
    return respJson.d.GetContextWebInformation.FormDigestValue;
  } catch(error) {
    console.error(error);
  }
};

async function saveFile(instance, folderUrl, fileName) {
  let annotationManager = instance.Core.annotationManager;
  const xfdfString = await annotationManager.exportAnnotations();
  const fileData = await instance.Core.documentViewer.getDocument().getFileData({ xfdfString });
  const digest = await getFormDigestValue();
  
  const fileArray = new Uint8Array(fileData);
  const file = new File([fileArray], fileName, {
    type: 'application/pdf'
  });
  const resp = await fetch(`${window.location.origin}/sites/5s4vrg/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files/add(url='${fileName}', overwrite=true)`, {
    method: 'POST',
    body: file,
    headers: {
      'accept': 'application/json; odata=verbose',
      'X-RequestDigest': digest,
      'Content-Length': file.byteLength
    }
  });
  const respJson = await resp.json();
}

function createSavedModal(instance) {
  const divInput = document.createElement('div');
  divInput.innerText = 'File saved successfully.';
  const modal = {
    dataElement: 'savedModal',
    body: {
      className: 'myCustomModal-body',
      style: {
        'text-align': 'center'
      },
      children: [divInput]
    }
  }
  instance.UI.addCustomModal(modal);
}