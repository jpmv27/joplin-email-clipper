'use strict';

ChromeUtils.import("resource:///modules/jsmime.jsm");
ChromeUtils.import("resource:///modules/gloda/mimemsg.js");

class JEC_Joplin {
  constructor() {
    this.port_ = -1;
  }

  async connect() {
    const MINPORT = 41184;
    const MAXPORT = MINPORT + 10;

    for (let port = MINPORT; port <= MAXPORT; port++) {
      try {
        let response = await this.request_({
          method: 'GET',
          url: 'http://127.0.0.1:' + port.toString() + '/ping'
        });
        if (response === 'JoplinClipperServer') {
          this.port_ = port;
          return true;
        }
      }
      catch (error) {
        // ignore
      }
    }
    this.port_ = -1;
    return false;
  }

  async createNote(subject, body, notebookId) {
    await this.request_({
      method: 'POST',
      url: 'http://127.0.0.1:' + this.port_.toString() + '/notes',
      headers: { 'Content-Type': 'application/json' },
      params: '{ "title": ' + JSON.stringify(subject) +
              ', "body": ' + JSON.stringify(body) +
              ', "parent_id": ' + JSON.stringify(notebookId) + ' }'
    });
  }

  get connected() {
    return (this.port_ !== -1);
  }

  async getNotebooks() {
    let folders = await this.request_({
      method: 'GET',
      url: 'http://127.0.0.1:' + this.port_.toString() + '/folders'
    });

    return JSON.parse(folders);
  }

  get port() {
    return this.port_;
  }

  request_(opts) {
    // https://stackoverflow.com/questions/30008114/how-do-i-promisify-native-xhr
    return new Promise(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      xhr.open(opts.method, opts.url);
      xhr.onload = function () {
        if (this.status >= 200 && this.status < 300) {
          resolve(xhr.response);
        }
        else {
          reject({
            status: this.status,
            statusText: xhr.statusText
          });
        }
      };
      xhr.onerror = function () {
        reject({
          status: this.status,
          statusText: xhr.statusText
        });
      };
      if (opts.headers) {
        Object.keys(opts.headers).forEach(function (key) {
          xhr.setRequestHeader(key, opts.headers[key]);
        });
      }
      var params = opts.params;
      // We'll need to stringify if we've been given an object
      // If we have a string, this is skipped.
      if (params && typeof params === 'object') {
        params = Object.keys(params).map(function (key) {
          return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
        }).join('&');
      }
      xhr.send(params);
    });
  }
}

class JEC_Popup {
  constructor() {
    this.window_ = null;
  }

  set body(val) {
    let b = this.window_.document.getElementById('joplin-preview-body');
    b.value = val;
    this.window_.sizeToContent();
  }

  close() {
    this.window_.close();
    this.window_ = null;
  }

  getConfirmation() {
    return new Promise((resolve) => {
      let c = this.window_.document.getElementById('joplin-confirm');
      c.disabled = false;
      c.addEventListener('command', function() {
        resolve(true);
      }, { once: true });
    });
  }

  get notebookId() {
    let menu = this.window_.document.getElementById('joplin-notebook-menu');
    return menu.selectedItem.getAttribute('value');
  }

  set notebooks(val) {
    // http://forums.mozillazine.org/viewtopic.php?t=1118715
    let menu = this.window_.document.getElementById('joplin-notebook-menu');
    let list = this.window_.document.getElementById('joplin-notebook-list');

    val.forEach((i) => {
      let menuItem = document.createElement('menuitem');
      menuItem.setAttribute('label', i.title);
      menuItem.setAttribute('value', i.id);
      list.appendChild(menuItem);
    });

    menu.selectedIndex = 0;
    menu.disabled = false;
  }

  open() {
    return new Promise((resolve) => {
      this.window_ = window.open(
        'chrome://emailclipper/content/popup.xul',
        'joplin',
        'chrome,resizable,centerscreen,scrollbars');
      this.window_.onload = function () {
        resolve(true);
      };
    });
  }

  set status(val) {
    let s = this.window_.document.getElementById('joplin-status');
    s.value = val;
  }

  set subject(val) {
    let s = this.window_.document.getElementById('joplin-preview-subject');
    s.value = val;
  }
}

class JEC_Message {
  constructor(header) {
    this.header_ = header;
    this.message_ = null;
  }

  get attachmentNames() {
    return this.message_.allUserAttachments.map(att => att.name);
  }

  get cc() {
    return jsmime.headerparser.decodeRFC2047Words(this.header_.ccList);
  }

  download() {
    return new Promise((resolve, reject) => {
      MsgHdrToMimeMessage(
        this.header_,
        this,
        function (hdr, msg) {
          if (msg) {
            this.message_ = msg;
            resolve(true);
          }
          else {
            reject('msg is null');
          }
        },
        true /* allowDownload */,
        { partsOnDemand: true, examineEncryptedParts: true });
    });
  }

  get from() {
    return jsmime.headerparser.decodeRFC2047Words(this.header_.author);
  }

  get plainBody() {
    return this.message_.coerceBodyToPlaintext(this.header_.folder);
  }

  get subject() {
    return jsmime.headerparser.decodeRFC2047Words(this.header_.subject);
  }

  get to() {
    return jsmime.headerparser.decodeRFC2047Words(this.header_.recipients);
  }
}

class JEC_Thunderbird {
  constructor() {
  }

  getCurrentMessage() {
    return new JEC_Message(gMessageDisplay.displayedMessage);
  }
}

class JEC_EmailClipper {
  constructor() {
    this.popup_ = new JEC_Popup();
    this.tbird_ = new JEC_Thunderbird();
    this.joplin_ = new JEC_Joplin();
  }

  async connectToJoplin_() {
    while (!this.joplin_.connected) {
      this.popup_.status = 'Looking for service';
      if (!await this.joplin_.connect()) {
        await this.sleep_(1000);
      }
    }

    this.popup_.status = 'Ready on port ' + this.joplin_.port.toString();
    return true;
  }

  flattenNotebookList_(list) {
    let notebooks = [];

    function helper(item, level) {
      notebooks.push({
        'title': ' . '.repeat(level) + item.title,
        'id': item.id
      });

      if (item.children) {
        item.children.forEach((i) => helper(i, level + 1));
      }
    }

    list.forEach((i) => helper(i, 0));

    return notebooks;
  }

  messageToNote_(msg) {
    const titleBlock = [
      { prop: 'from',    label: 'From',    optional: false },
      { prop: 'subject', label: 'Subject', optional: false },
      { prop: 'to',      label: 'To',      optional: false },
      { prop: 'cc',      label: 'Cc',      optional: true  }
    ];

    // Find maximum length of labels and properties
    let maxLabel = 0;
    let maxProp = 0;
    titleBlock.forEach((i) => {
      if (msg[i.prop]) {
        maxLabel = Math.max(maxLabel, i.label.length);
        maxProp = Math.max(maxProp, msg[i.prop].length);
      }
    });

    let note = '| ' + ' '.repeat(maxLabel + 5) + ' | ' + ' '.repeat(maxProp) + ' |\n' +
               '| ' + '-'.repeat(maxLabel + 5) + ' | ' + '-'.repeat(maxProp) + ' |\n';

    titleBlock.forEach((i) => {
      if (!i.optional || msg[i.prop]) {
        note += '| **' + i.label + ':**' + ' '.repeat(maxLabel - i.label.length) +
          ' | ' + msg[i.prop] + ' '.repeat(maxProp - msg[i.prop].length) + ' |\n';
      }
    });

    note += '\n' + msg.plainBody;

    return note;
  }

  async sendToJoplin() {
    await this.popup_.open();

    let msg = this.tbird_.getCurrentMessage();

    this.popup_.subject = msg.subject;
    this.popup_.body = 'Downloading...';
    await msg.download();

    let note = this.messageToNote_(msg);
    this.popup_.body = note;

    await this.connectToJoplin_();
    this.popup_.notebooks = this.flattenNotebookList_(await this.joplin_.getNotebooks());

    await this.popup_.getConfirmation();

    await this.joplin_.createNote(msg.subject, note, this.popup_.notebookId);

    this.popup_.close();
  }

  async sleep_(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }
}

/* exported JEC_sendToJoplin */
function JEC_sendToJoplin() {
	console.info('sendToJoplin started');
  let clipper = new JEC_EmailClipper();
  clipper.sendToJoplin()
    .then(() => { console.info('sendToJoplin done'); })
    .catch((error) => { console.error('sendToJoplin failed: ' + error.toString()) });
}
