console.log('instantiating webviewer');

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