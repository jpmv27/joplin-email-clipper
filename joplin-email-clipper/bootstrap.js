'use strict';

let { JEC_EmailClipper } = ChromeUtils.import('chrome://emailclipper/content/emailclipper.jsm');

let { Services } = ChromeUtils.import('resource://gre/modules/Services.jsm');

let onLoadObserver = {
    observe: function(aSubject, aTopic, aData) {
        let mainWindow = Services.wm.getMostRecentWindow('mail:3pane');
        if (mainWindow) {
						const elem = mainWindow.document.getElementById('viewSourceMenuItem');
						const menuItem = mainWindow.document.createElement('menuitem');
						menuItem.setAttribute('label', 'Send to Joplin');
						menuItem.setAttribute('accesskey', 'J');
						elem.parentNode.insertBefore(menuItem, elem);
						console.log('clipper success');
        } else {
            console.log('clipper failed');
        }
    }
}

function startup(data, reason) {
	console.log('clipper startup ' + reason);
  Services.obs.addObserver(onLoadObserver, "mail-startup-done", false);
}

function shutdown(data, reason) {
	console.log('clipper shutdown ' + reason);
}

function install(data, reason) {
  // Ignore
}

function uninstall(data, reason) {
  // Ignore
}
