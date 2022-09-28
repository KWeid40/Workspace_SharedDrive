var ss = SpreadsheetApp.openByUrl('Sheet URL Here');
var sheet = ss.getSheetByName('Active sheet here');

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Drive Functions')
    .addItem('Clear','setHeaders')
    .addItem('Get Drives','listSharedDrives')
    .addToUi();
}

function setHeaders(){
  sheet.clearContents();
  sheet.appendRow(['Drive Name','User','User permission']);
}

function listSharedDrives() {
  let pageToken = null;
  const response = [];
  let sharedDrives = Drive.Drives.list({
          maxResults: 5,
          useDomainAdminAccess: true,
          hidden: false,
        });
     
  
    while(sharedDrives.nextPageToken){
     sharedDrives = Drive.Drives.list({
          maxResults: 50,
          useDomainAdminAccess: true,
          hidden: false,
          pageToken: pageToken,
        });
    let sharedDrivesItems = sharedDrives.items;
    //console.log(sharedDrivesItems+'12221');


    sharedDrivesItems.forEach((item) => {
    //console.log(item);
    let driveID = item.id;
    //console.log(item.name);
    getUsers(driveID, item.name);
    
    });
    console.log(sharedDrives);
    pageToken = sharedDrives.nextPageToken;
    }
  

  console.log(response +'111');
}

function getUsers(drivID, driveName){
  const args = {
    supportsAllDrives: true,
    useDomainAdminAccess: true,
  };
      let pList = Drive.Permissions.list(drivID, args);
    let editors = pList.items;
    console.log(editors.length);
    for (var i = 0; i < editors.length; i++) {
      let email = editors[i].emailAddress;
      let role = editors[i].role;
      addRow(driveName, email, role)
    }
    if(editors.length == 0){
      addRow(driveName,"No users","No permissions")
    }
  
}

function addRow(driveName, user, permission){

  sheet.appendRow([driveName,user,permission]);
}
