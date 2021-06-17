///////// Set your properties /////////////////
const sharedSpreadId = '1NT9H4uQ7Jyzn0df7ZZPNtHuXM3Sf_Lm6eGFNxRa1wPI'
const sharedFolderId = '1derhMC8Aoq8ULDoZ0Vde94_e3XjzJGlo'
///////////////////////////////////////////////////

const driveFolder = DriveApp.getFolderById(sharedFolderId)
const spreadsheet = SpreadsheetApp.openById(sharedSpreadId)

const fileIdRow = 0
const copyIdRow = 1
const emailRow = 3
const expirationRow = 5

///////////////// Trigger ////////////////////////////
function removeExpiredFiles() {
    var todayTime = new Date().getTime()
    var range = spreadsheet.getDataRange()
    var sheet = range.getValues()
    var values = sheet.slice(1, sheet.length)
    var expiredRows = []
    var count = 0
    values.forEach(function (row, index) {
        var copyId = row[copyIdRow]
        var expiration = row[expirationRow]
        var expirationDate = new Date(expiration)
        // If expired Remove Copy file
        if (expirationDate.getTime() < todayTime) {
            // Remove File
            let expiredFile = DriveApp.getFileById(copyId)
            expiredFile.setTrashed(true)
            // Add expired row
            expiredRows.push(index)
        }
    })
    expiredRows.forEach(function (index) {
        spreadsheet.deleteRow(index + 2 - count)
        count += 1
    })
}

/////////////////// Spreadsheet ///////////////////////
function getVersions(myFileId: string, myEmails: string[]) {
    var range = spreadsheet.getDataRange()

    var sheet = range.getValues()

    var values = sheet.slice(1, sheet.length)

    var versions = myEmails.map((_) => 0)
    values.forEach((row) => {
        let fileId = row[fileIdRow]
        let emails = (row[emailRow] as string).split(',')
        if (fileId == myFileId) {
            myEmails.forEach((myEmail, i) => {
                if (emails.indexOf(myEmail) > -1) versions[i] += 1
            })
        }
    })

    return versions
}
function getViewers(fileId: string) {
    const file = DriveApp.getFileById(fileId)
    const viewers = file.getViewers()
    return viewers.map((viewer) => {
        return viewer.getEmail()
    })
}

function modifySpread(contents: string[]) {
    spreadsheet.appendRow(contents)
}

////////////////// Share File ////////////////////
const customMessage = (
    extraMessage: string,
    expiration: string,
    fileName: string,
    version: number = 0
) => {
    var verString = version == 0 ? '' : `_Ver${version}`
    return `${extraMessage}
    ファイル名: ${fileName}${verString}
    このファイルは${expiration}には削除しますので、必要に応じてそれまでにダウンロードいただくようお願いします。`
}

function uploadFromDrive(e: any) {
    // Read Form inputs
    var arr = e['array[]'] as string[] | string
    var ids = e.fileId as string
    var expiration = e.expiration as string
    var message = e.message as string
    // Get Files ids and emails
    var emails = arr.toString().split(',')
    var fileIds = ids.split(',')

    if (fileIds.length == 1) {
        var fileId = fileIds[0]
        try {
            // is file
            var file = DriveApp.getFileById(fileId)
            var copy = file.makeCopy(file.getName(), driveFolder)
            var versions = getVersions(fileId, emails)

            shareFileToUsers(
                emails,
                fileId,
                copy.getId(),
                file.getName(),
                expiration,
                versions,
                message
            )
        } catch (error) {
            // is folder
            var { id: folderId, name: folderName } = copySingleFolder(
                fileIds[0]
            )
            shareFolderToUsers(
                emails,
                folderId,
                folderName,
                expiration,
                message
            )
        }
    } else {
        // Find folder where files are
        var { id: folderId, name: folderName } = copyMultipleFiles(fileIds)
        shareFolderToUsers(emails, folderId, folderName, expiration, message)
    }
}

function getFiles(
    rootFolder: GoogleAppsScript.Drive.Folder,
    destFolder: GoogleAppsScript.Drive.Folder
) {
    var files = rootFolder.getFiles()
    while (files.hasNext()) {
        var file = files.next()
        file.makeCopy(destFolder)
    }
    var folders = rootFolder.getFolders()
    while (folders.hasNext()) {
        var folder = folders.next()
        var copyFolder = destFolder.createFolder(folder.getName())
        getFiles(folder, copyFolder)
    }
}
function copySingleFolder(id: string) {
    var folder = DriveApp.getFolderById(id)
    var outFolder = driveFolder.createFolder(folder.getName())
    getFiles(folder, outFolder)

    return { id: outFolder.getId(), name: outFolder.getName() }
}

function copyMultipleFiles(ids: string[]) {
    var names: Array<string> = []
    const tempName = Math.random().toString(36).substr(2, 9)
    const outFolder = driveFolder.createFolder(tempName)
    ids.forEach((id) => {
        try {
            let file = DriveApp.getFileById(id)
            file.makeCopy(file.getName(), outFolder)
            names.push(file.getName())
        } catch (err) {
            var folder = DriveApp.getFolderById(id)
            var copyFolder = outFolder.createFolder(folder.getName())
            getFiles(folder, copyFolder)
            names.push(folder.getName())
        }
    })
    outFolder.setName(names.join('__'))

    return { id: outFolder.getId(), name: outFolder.getName() }
}

function shareFileToUsers(
    emails: string[],
    fileId: string,
    copyId: string,
    fileName: string,
    expiration: string,
    versions: number[],
    extraMessage: string
) {
    var date = new Date(expiration)
    date.setHours(date.getHours() + 1)

    emails.forEach((email, i) => {
        let body = customMessage(
            extraMessage,
            expiration,
            fileName,
            versions[i]
        )
        shareFileToUser(email, fileId, body, date)
    })

    modifySpread([
        fileId,
        copyId,
        fileName,
        emails.toString(),
        extraMessage,
        date.toISOString(),
        new Date().toTimeString(),
    ])
}

function shareFolderToUsers(
    emails: string[],
    fileId: string,
    fileName: string,
    expiration: string,
    extraMessage: string
) {
    var date = new Date(expiration)
    date.setHours(date.getHours() + 1)

    emails.forEach((email) => {
        let body = customMessage(extraMessage, expiration, fileName)
        shareFileToUser(email, fileId, body, date)
    })

    modifySpread([
        fileId,
        fileId,
        fileName,
        emails.toString(),
        extraMessage,
        date.toISOString(),
        new Date().toTimeString(),
    ])
}

function shareFileToUser(
    email: string,
    fileId: string,
    body: string,
    expiration: Date
) {
    var permission = Drive.Permissions.insert(
        {
            value: email,
            type: 'user',
            role: 'reader',
            withLink: false,
        },
        fileId,
        {
            sendNotificationEmails: true,
            emailMessage: body,
        }
    )
    Drive.Permissions.patch(
        {
            expirationDate: expiration.toISOString(),
        },
        fileId,
        permission.id
    )
}

/////////// Google Drive Picker /////////////////////////

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Picker')
        .addItem('Start', 'showPicker')
        .addToUi()
}

function showPicker() {
    var html = HtmlService.createHtmlOutputFromFile('dialog.html')
        .setWidth(600)
        .setHeight(425)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    SpreadsheetApp.getUi().showModalDialog(html, 'Select a file')
}

function getOAuthToken() {
    DriveApp.getRootFolder()
    return ScriptApp.getOAuthToken()
}

///////////////// DO GET ///////////////////////////////////
function doGet() {
    return HtmlService.createTemplateFromFile('form.html')
        .evaluate()
        .setTitle('Google Visitor Manager')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
}
