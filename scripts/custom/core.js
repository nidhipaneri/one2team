
(function () {
    "use strict";

    One2Teamapp.AppViewModel = (function () {

        var docs = ko.observableArray();
        var breadCrumbs = ko.observableArray();
        var keepGoing = true;
        var timerid;
        var tempfolderUrl = null;
        var breadcrumbIndex = null;
        var addBreadcrumb = null;
        var isTokenRefreshed = false;

        setTimeout(function () {
            keepGoing = false;
        }, One2Teamapp.Param.timeOutToGetAccessToken);

        function GetToken() {
            try {
                if (One2Teamapp.Param.domain &&
                One2Teamapp.Param.appWebUrl &&
                One2Teamapp.Param.documentUrl &&
                One2Teamapp.Param.applicationId &&
                One2Teamapp.Param.clientsecret &&
                One2Teamapp.Param.redirectUrl) {
                    keepGoing = true;
                    timerid = setTimeout(CheckToken, 1000);
                    window.open(One2Teamapp.Param.msLoginUrl + One2Teamapp.Param.domain + '/oauth2/authorize?resource=' + One2Teamapp.Param.appWebUrl + '&client_id=' + One2Teamapp.Param.applicationId + '&response_type=id_token+code&redirect_uri=' + One2Teamapp.Param.redirectUrl + '&response_mode=form_post&scope=openid&state=12345&nonce=7362CAEA-9CA5-4B43-9BA3-34D7C303EBA7');
                }
                else {
                    One2Teamapp.Utility.RestApiCall(One2Teamapp.Param.urlToGetconfigParam, "GET", One2Teamapp.Param.headerToGetconfigParam, function (callback) {
                        if (callback.status && callback.status == 'error') {
                            alert(One2Teamapp.Utility.customMessage.ConfigParamError);
                        }
                        else {
                            One2Teamapp.Param.domain = callback.domain;
                            One2Teamapp.Param.appWebUrl = callback.appWebUrl;
                            One2Teamapp.Param.documentUrl = callback.documentUrl;
                            One2Teamapp.Param.applicationId = callback.applicationId;
                            One2Teamapp.Param.clientsecret = callback.clientSecret;
                            One2Teamapp.Param.redirectUrl = callback.redirectUrl;

                            keepGoing = true;
                            timerid = setTimeout(CheckToken, 1000);
                            window.open(One2Teamapp.Param.msLoginUrl + One2Teamapp.Param.domain + '/oauth2/authorize?resource=' + One2Teamapp.Param.appWebUrl + '&client_id=' + One2Teamapp.Param.applicationId + '&response_type=id_token+code&redirect_uri=' + One2Teamapp.Param.redirectUrl + '&response_mode=form_post&scope=openid&state=12345&nonce=7362CAEA-9CA5-4B43-9BA3-34D7C303EBA7');
                        }
                    });
                }
            }
            catch (e) {
                alert(One2Teamapp.Utility.customMessage.ConfigParamError);
            }
        }

        function GetRefreshToken() {
            localStorage.setItem("accesstoken", null);
            isTokenRefreshed = true;
            keepGoing = true;
            timerid = setTimeout(CheckToken, 1000);
            window.open('https://login.microsoftonline.com/' + One2Teamapp.Param.domain + '/oauth2/authorize?resource=' + One2Teamapp.Param.appWebUrl + '&client_id=' + One2Teamapp.Param.applicationId + '&response_type=id_token+code&redirect_uri=' + One2Teamapp.Param.redirectUrl + '&response_mode=form_post&refresh_token=' + localStorage.getItem("refreshtoken") + '&scope=openid&state=12345&nonce=7362CAEA-9CA5-4B43-9BA3-34D7C303EBA7');
        }

        function Logout() {
            localStorage.setItem("accesstoken", null);
            window.open('https://login.microsoftonline.com/' + One2Teamapp.Param.domain + '/oauth2/logout');
        }

        function UploadToSharepoint() {
            tempfolderUrl = breadCrumbs()[breadCrumbs().length - 1].url;
            window.open(One2Teamapp.Param.appWebUrl + One2Teamapp.Param.documentUrl + '/Forms/AllItems.aspx?RootFolder=' + tempfolderUrl);
        }

        function UploadToSPhere() {
            tempfolderUrl = breadCrumbs()[breadCrumbs().length - 1].url;
            One2Teamapp.Utility.UploadDocument(tempfolderUrl);
        }

        function CreateNewFolder() {
            var folderName = prompt(One2Teamapp.Utility.customMessage.enterFolerName, "");
            if (folderName.trim()) {
                tempfolderUrl = breadCrumbs()[breadCrumbs().length - 1].url;
                var url = "/_api/web/getfolderbyserverrelativeurl('" + tempfolderUrl + "')/Folders/Add('" + folderName.trim() + "')";
                var header = {
                    "Authorization": "Bearer " + localStorage.getItem("accesstoken"),
                    "accept": "application/json;odata=verbose",
                    "content-type": "application/json; odata=verbose",
                    "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
                };
                One2Teamapp.Utility.RestApiCall(url, "POST", header, function (callback) {
                    if (callback.status && callback.status == 'error') {
                        alert(One2Teamapp.Utility.customMessage.newFolderError + callback.errormessage);
                    }
                    else {
                        GetItems(tempfolderUrl);
                    }
                });
            }
            else {
                alert(One2Teamapp.Utility.customMessage.folderNameError);
                return false;
            }
        }

        function CheckToken() {
            if (localStorage.getItem("accesstoken") != "" && localStorage.getItem("accesstoken") != null && localStorage.getItem("accesstoken") != "null") {
                AbortTimer();
                if (tempfolderUrl != null) {
                    if (breadcrumbIndex != null) {
                        breadCrumbs.splice(item.index);
                        breadcrumbIndex = null;
                    }
                    if (addBreadcrumb != null) {
                        AddCrumb(addBreadcrumb.ServerRelativeUrl, addBreadcrumb.Name, true, breadCrumbs().length + 1);
                        addBreadcrumb = null;
                    }
                    GetItems(tempfolderUrl);
                    tempfolderUrl = null;
                }
                else {
                    GetFoldersAndFiles();
                }
            }
            else if (keepGoing)
                timerid = setTimeout(CheckToken, 1000); // repeat
            else {
                $("#wait").css("display", "none");
                alert(One2Teamapp.Utility.customMessage.docLibOpenError);
                AbortTimer();
            }
        }

        function AbortTimer() {
            clearTimeout(timerid);
        }

        function OnFolderClick(item) {
            if (localStorage.getItem("accesstoken") != null && localStorage.getItem("accesstoken") != "null" && localStorage.getItem("accesstoken") != "") {
                AddCrumb(item.ServerRelativeUrl, item.Name, true, breadCrumbs().length + 1);
                GetItems(item.ServerRelativeUrl);
            }
            else {
                tempfolderUrl = item.url;
                addBreadcrumb = item;
                GetToken();
            }
        }

        function OnBreadCrumbClick(item) {
            if (localStorage.getItem("accesstoken") != null && localStorage.getItem("accesstoken") != "null" && localStorage.getItem("accesstoken") != "") {
                breadCrumbs.splice(item.index);
                GetItems(item.url);
            }
            else {
                tempfolderUrl = item.url;
                breadcrumbIndex = item.index;
                GetToken();
            }
        }

        function AddCrumb(url, text, active, index) {
            breadCrumbs.push({
                'url': url, 'text': text, 'active': active, 'index': index
            });
        }

        function GetFoldersAndFiles() {
            if (localStorage.getItem("accesstoken") != null && localStorage.getItem("accesstoken") != "null" && localStorage.getItem("accesstoken") != "") {
                $("#wait").css("display", "block");
                breadCrumbs([]);
                AddCrumb(One2Teamapp.Param.documentUrl, "Documents", true, 1);
                GetItems(One2Teamapp.Param.documentUrl);
                $("#docModal").modal('show');
            }
            else {
                GetToken();
            }
        }

        function GetItems(folderUrl) {
            tempfolderUrl = null;
            if (localStorage.getItem("accesstoken") != null && localStorage.getItem("accesstoken") != "null" && localStorage.getItem("accesstoken") != "") {
                $("#wait").css("display", "block");

                var header = {
                    "Authorization": "Bearer " + localStorage.getItem("accesstoken"),
                    "Access-Control-Allow-Origin": "*",
                    "accept": "application/json;odata=verbose"
                };
                var url = "/_api/web/GetFolderByServerRelativeUrl('" + folderUrl + "')?$expand=Folders,Files";
                One2Teamapp.Utility.RestApiCall(url, "GET", header, function (callback) {
                    if (callback.status && callback.status == 'error') {
                        if (!isTokenRefreshed)
                            GetRefreshToken();
                        else {
                            $("#wait").css("display", "none");
                            alert(One2Teamapp.Utility.customMessage.docLibOpenError + callback.errormessage);
                            isTokenRefreshed = false;
                        }
                    }
                    else {
                        docs([]);
                        $.each(callback.Folders.results, function (obj, item) {
                            var folder = {
                                Name: item.Name,
                                IsFolder: true,
                                URL: One2Teamapp.Param.appWebUrl + "/" + item.ServerRelativeUrl,
                                ServerRelativeUrl: item.ServerRelativeUrl,
                                isSelected: false,
                                modified: item.TimeLastModified,
                                icon: One2Teamapp.Utility.GetIcon(item.Name.substr(item.Name.lastIndexOf('.') + 1))
                            };
                            docs.push(folder);
                        });
                        $.each(callback.Files.results, function (obj, item) {
                            var file = {
                                Name: item.Name,
                                IsFolder: false,
                                URL: item.LinkingUri == null ? One2Teamapp.Param.appWebUrl + "/" + item.ServerRelativeUrl : item.LinkingUri,
                                ServerRelativeUrl: item.ServerRelativeUrl,
                                isSelected: false,
                                modified: item.TimeLastModified,
                                icon: One2Teamapp.Utility.GetIcon(item.Name.substr(item.Name.lastIndexOf('.') + 1))
                            };
                            docs.push(file);
                        });
                        // sort by date modified using moment.js
                        docs.sort(function (left, right) {
                            return moment.utc(left.modified).diff(moment.utc(right.modified))
                        }).reverse();
                        // sort by IsFolder: all folders on top
                        docs.sort(function (x, y) {
                            return (x.IsFolder === y.IsFolder) ? 0 : x.IsFolder ? -1 : 1;
                        });
                        $("#wait").css("display", "none");
                    }
                });
            }
            else {
                tempfolderUrl = folderUrl;
                GetToken();
            }
        }

        function AddToList() {
            var docsToAdd = [];
            $.each(docs(), function (obj, item) {
                if (item.isSelected)
                    docsToAdd.push({
                        'Name': item.Name, 'URL': item.URL
                    });
            });
            //This JSON will be returned to One2Team page
            alert(JSON.stringify(docsToAdd));
        }

        function Init() {
            try {
                One2Teamapp.Utility.RestApiCall(One2Teamapp.Param.urlToGetconfigParam, "GET", One2Teamapp.Param.headerToGetconfigParam, function (callback) {
                    if (callback.status && callback.status == 'error') {
                        alert(One2Teamapp.Utility.customMessage.ConfigParamError);
                    }
                    else {
                        One2Teamapp.Param.domain = callback.domain;
                        One2Teamapp.Param.appWebUrl = callback.appWebUrl;
                        One2Teamapp.Param.documentUrl = callback.documentUrl;
                        One2Teamapp.Param.applicationId = callback.applicationId;
                        One2Teamapp.Param.clientsecret = callback.clientSecret;
                        One2Teamapp.Param.redirectUrl = callback.redirectUrl;
                    }
                });
            }
            catch (e) {
                alert(One2Teamapp.Utility.customMessage.ConfigParamError);
            }
        }

        document.addEventListener('visibilitychange', function () {
            if (document.visibilityState === 'visible' && tempfolderUrl != null && localStorage.getItem("accesstoken") != null && localStorage.getItem("accesstoken") != "null" && localStorage.getItem("accesstoken") != "") {
                GetItems(tempfolderUrl);
            }
        });

        $("html").on("dragover", function (e) {
            e.preventDefault();
            e.stopPropagation();
            $("#uploadfile").addClass('upload-area-highlight');
        });

        $("html").on("drop", function (e) {
            e.preventDefault();
            e.stopPropagation();
            $("#uploadfile").removeClass('upload-area-highlight');
        });

        $('.modal-dialog').on('mouseout', function (e) {
            e.stopPropagation();
            e.preventDefault();
            $("#uploadfile").removeClass('upload-area-highlight');
        });

        $('.upload-area').on('dragenter', function (e) {
            e.stopPropagation();
            e.preventDefault();
            $("#uploadfile").addClass('upload-area-highlight');
        });

        $('.upload-area').on('dragover', function (e) {
            e.stopPropagation();
            e.preventDefault();
            $("#uploadfile").addClass('upload-area-highlight');
        });

        $('.upload-area').on('drop', function (e) {
            e.stopPropagation();
            e.preventDefault();
            $("#uploadfile").removeClass('upload-area-highlight');
            $("#wait").css("display", "block");
            var filesToUpload = e.originalEvent.dataTransfer.files;
            var currentFolder = One2Teamapp.AppViewModel.breadCrumbs()[One2Teamapp.AppViewModel.breadCrumbs().length - 1].url; // "/One2TeamDocuments/One2Team Docs/Images";

            One2Teamapp.Utility.ReadAndUploadFiles(filesToUpload, currentFolder);
        });

        Init();

        return {
            docs: docs,
            breadCrumbs: breadCrumbs,
            GetItems: GetItems,
            OnFolderClick: OnFolderClick,
            OnBreadCrumbClick: OnBreadCrumbClick,
            GetToken: GetToken,
            GetRefreshToken: GetRefreshToken,
            GetFoldersAndFiles: GetFoldersAndFiles,
            AddToList: AddToList,
            Logout: Logout,
            UploadToSharepoint: UploadToSharepoint,
            UploadToSPhere: UploadToSPhere,
            CreateNewFolder: CreateNewFolder
        };
    })();
    ko.applyBindings(One2Teamapp.AppViewModel);
})();