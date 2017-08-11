var One2Teamapp = One2Teamapp || {};

One2Teamapp.Param = (function () {

    var msLoginUrl = "https://login.microsoftonline.com/";
    var urlToGetconfigParam = "https://api-devel.one2team.com/api/helpers/o365/settings";
    var headerToGetconfigParam = {
        "Authorization": "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhdWQiOiJ3eURmVTY3aGxJN1Z4UTc5UWRlakkzYktQbkVXNTV4RSIsImV4cCI6IjE1MDg4MzgzMjgiLCJpYXQiOiIxNTAxMDYyMzI4Iiwic3ViIjoiYXV0aDB8a2hhbGlmZTItbW9iaWxlL2FsdCIsInNjb3BlcyI6W119.8C5bkFog8tQPrpVswKFmpWp+G2wBkxN1k49iOYOiSQs="
    };

    var appWebUrl = "";
    var documentUrl = "";
    var applicationId = "";
    var clientsecret = "";
    var redirectUrl = "";
    var domain = "";


    //var appWebUrl = "https://advaiya.sharepoint.com";
    //var documentUrl = "/One2TeamDocuments";
    //var applicationId = "61d6e529-199a-402a-9da0-3f1300c50d61";
    //var clientsecret = "FCsCCcGjkI8cM0LsVpiw9OpGTwTIfI71MIiMvOuMqrM=";
    //var redirectUrl = "http://one2teamappintegration.azurewebsites.net/oauth.php";
    //var domain = "advaiya.com";

    var timeOutToGetAccessToken = 120000;

    return {
        appWebUrl: appWebUrl,
        documentUrl: documentUrl,
        applicationId: applicationId,
        clientsecret: clientsecret,
        redirectUrl: redirectUrl,
        domain: domain,
        urlToGetconfigParam: urlToGetconfigParam,
        headerToGetconfigParam: headerToGetconfigParam,
        timeOutToGetAccessToken: timeOutToGetAccessToken,
        msLoginUrl: msLoginUrl
    };

})();

One2Teamapp.Utility = (function () {
    'use strict';

    var customMessage = {
        "ConfigParamError": "Unable to get Configuration parameters to proceed further with O365 Integration! ",
        "newFolderError": "Unable to create a new folder, please try again! ",
        "docLibOpenError": "Unable to access Office 365 Document library, please try again! ",
        "noFolderUpload": "Folder cannot be uploaded! ",
        "fileUploadError": "A error occurred during file upload! ",
        "enterFolerName": "Enter Folder Name",
        "folderNameError": "Folder name cannot be blank! "
    };

    function RestApiCall(url, type, header, callback) {
        $.ajax({
            url: One2Teamapp.Param.appWebUrl + url,
            type: type,
            crossDomain: true,
            headers: header,
            success: function (data) {
                if (data.d)
                    callback(data.d);
                else
                    callback(data);
            },
            error: function (xhr, statusText, err) {
                var message = "";
                if (xhr.responseText) {
                    var response = JSON.parse(xhr.responseText);
                    message = response ? response.error.message.value : statusText;
                }
                callback({ status: "error", errormessage: message });
            }
        });
    }

    function GetIcon(fileExtension) {
        switch (fileExtension) {
            case 'pdf': return 'images/icons/pdf.png';
            case 'pst': case 'ost': return 'images/icons/outlook.png';
            case 'vsd': case 'ost': return 'images/icons/visio.png';
            case 'zip': return "images/icons/zip.png";
            case 'rar': return "images/icons/rar.png";
            case 'gz': return "images/icons/zip.png";
            case 'swf': return "images/icons/swf.jpg";
            case 'odt': return "images/icons/odt.png";
            case 'wbk': case 'wiz': case 'doc': case 'dot': case 'docx': return 'images/icons/word.png';
            case 'slk': case 'xla': case 'xlam': case 'xlc': case 'xld': case 'xlk':
            case 'xll': case 'xlm': case 'xls': case 'xlsb': case 'xlsm': case 'xlt':
            case 'xltm': case 'xlw': case 'xlsx': return 'images/icons/excel.png';
            case 'pot': case 'potm': case 'ppa': case 'ppam': case 'pps': case 'ppsm':
            case 'ppt': case 'pptm': case 'pwz': case 'ppsx': case 'pptx': return 'images/icons/powerpoint.png';
            case 'ai': case 'eps': case 'ps': return "images/icons/ps.jpg";
            case 'jar': return "images/icons/jar.jpg";
            case 'rtf': return "images/icons/rtf.jpg";
            case 'mp3': return "images/icons/mp3.png";
            case 'jpe': case 'jpg': return 'images/icons/jpg.png';
            case 'jpeg': return 'images/icons/jpeg.png';
            case 'jpg': return "images/icons/jpg.png";
            case 'psd': return "images/icons/psd.png";
            case 'pnz': case 'png': return 'images/icons/png.png';
            case 'gif': return 'images/icons/gif.png';
            case 'tif': case 'tiff': return "images/icons/tiff.png";
            case 'bmp': return 'images/icons/bmp.png';
            case 'mpeg': return 'images/icons/mpeg.png';
            case 'mpa': case 'mpe': case 'mpg': return "images/icons/mpg.png";
            case 'mp4': case 'mp4v': return "images/icons/mp4.png";
            case 'mpv2': case 'mp2': case 'mp2v': case 'm1v': case 'm2v': case 'mov': case 'mqv': return "images/icons/mov.jpg";
            case 'htm': case 'html': case 'hxt': case 'shtml': return "images/icons/html.jpg";
            default: return 'images/icons/file.png';
        }
    }

    function ReadAndUploadFiles(files, serverRelativeUrlToFolder) {
        $("#wait").css("display", "block");
        var fileCount = files.length;
        var filesUploaded = 0;

        for (var i = 0; i < fileCount; i++) {
            if (!files[i].type) {
                alert(customMessage.noFolderUpload);
                filesUploaded++;
            }
            else {
                var getFile = getFileBuffer(i);
                getFile.done(function (arrayBuffer, i) {
                    var addFile = addFileToFolder(arrayBuffer, i);
                    addFile.done(function (file, status, xhr) {
                        filesUploaded++;
                        if (fileCount == filesUploaded) {
                            One2Teamapp.AppViewModel.GetItems(serverRelativeUrlToFolder);
                            filesUploaded = 0;
                            $("#wait").css("display", "none");
                        }
                    });
                    addFile.fail(function (jqXHR, textStatus, errorThrown) {
                        filesUploaded++;
                        if (fileCount == filesUploaded)
                            $("#wait").css("display", "none");
                        FailHandler(jqXHR, textStatus, errorThrown);
                    });
                });
                getFile.fail(function (jqXHR, textStatus, errorThrown, file) {
                    filesUploaded++;
                    if (fileCount == filesUploaded)
                        $("#wait").css("display", "none");
                    FailHandler(jqXHR, textStatus, errorThrown);
                });
            }
        }


        function getFileBuffer(i) {
            var deferred = jQuery.Deferred();
            var reader = new FileReader();
            reader.onloadend = function (e) {
                deferred.resolve(e.target.result, i);
            }
            reader.onerror = function (e) {
                deferred.reject(e.target.error);
            }
            reader.readAsArrayBuffer(files[i]);
            return deferred.promise();
        }


        function addFileToFolder(arrayBuffer, i) {
            var index = i;
            var fileName = files[index].name;

            var fileCollectionEndpoint = One2Teamapp.Param.appWebUrl + "/_api/web/getfolderbyserverrelativeurl('" + serverRelativeUrlToFolder + "')/Files/Add(url='" + fileName + "', overwrite=true)";
            return jQuery.ajax({
                url: fileCollectionEndpoint,
                type: "POST",
                data: arrayBuffer,
                processData: false,
                headers: {
                    "Authorization": "Bearer " + localStorage.getItem("accesstoken"),
                    "accept": "application/json;odata=verbose",
                    "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                    "content-length": arrayBuffer.byteLength
                }
            });
        }
    }

    function UploadDocument(currentFolder) {
        $("#wait").css("display", "block");
        var fileInput = jQuery('#file-3');
        ReadAndUploadFiles(fileInput[0].files, currentFolder);
    }

    //Error handling
    function FailHandler(jqXHR, textStatus, errorThrown) {

        var message = "";
        if (jqXHR.message) {
            message = jqXHR.message;
        }
        if (jqXHR.responseText) {
            var response = JSON.parse(jqXHR.responseText);
            message = response ? response.error.message.value : textStatus;
        }
        alert(customMessage.fileUploadError + message);
    }

    return {
        RestApiCall: RestApiCall,
        GetIcon: GetIcon,
        UploadDocument: UploadDocument,
        ReadAndUploadFiles: ReadAndUploadFiles,
        customMessage: customMessage
    };

})();

$(document).ready(function () {
    jQuery.support.cors = true;
    $('.modal').draggable();
});


