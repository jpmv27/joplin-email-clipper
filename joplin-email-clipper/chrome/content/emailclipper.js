'use strict';

ChromeUtils.import("resource:///modules/jsmime.jsm");
ChromeUtils.import("resource:///modules/gloda/mimemsg.js");

let lastSelectedNotebookId;

class JEC_JoplinError {
  constructor(status, statusText) {
    this.status = status;
    this.statusText = statusText;
  }

  toString() {
    return 'Joplin error (' + this.status.toString() + '): ' + this.statusText;
  }
}

class JEC_Joplin {
  constructor() {
    this.port_ = -1;
  }

  async connect() {
    const MINPORT = 41184;
    const MAXPORT = MINPORT + 10;

    for (let port = MINPORT; port <= MAXPORT; port++) {
      try {
        const response = await this.request_({
          method: 'GET',
          url: 'http://127.0.0.1:' + port.toString() + '/ping',
          timeout: 10000
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

  async createNote(subject, body, notebookId, tagIds) {
    const response = await this.request_({
      method: 'POST',
      url: 'http://127.0.0.1:' + this.port_.toString() + '/notes',
      headers: { 'Content-Type': 'application/json' },
      params: '{ "title": ' + JSON.stringify(subject) +
              ', "body": ' + JSON.stringify(body) +
              ', "parent_id": ' + JSON.stringify(notebookId) + ' }',
      timeout: 10000
    });

    const note = JSON.parse(response);

    await this.tagNote_(note.id, tagIds);
  }

  get connected() {
    return (this.port_ !== -1);
  }

  async getNotebooks() {
    const response = await this.request_({
      method: 'GET',
      url: 'http://127.0.0.1:' + this.port_.toString() + '/folders',
      timeout: 10000
    });

    return JSON.parse(response);
  }

  async getTags() {
    const response = await this.request_({
      method: 'GET',
      url: 'http://127.0.0.1:' + this.port_.toString() + '/tags',
      timeout: 10000
    });

    const rawTags = JSON.parse(response);

    // Remove duplicate tags. Keep one first seen in the list.
    const uniqueTags = rawTags.filter((element, index, array) =>
      index === array.findIndex(e => e.title === element.title));

    return uniqueTags.sort((a, b) => a.title.localeCompare(b.title));
  }

  get port() {
    return this.port_;
  }

  request_(opts) {
    return new Promise(function (resolve, reject) {
      const xhr = new XMLHttpRequest();
      xhr.open(opts.method, opts.url);

      xhr.onload = () => {
        if (xhr.status >= 200 && xhr.status < 300) {
          resolve(xhr.response);
        }
        else {
          reject(new JEC_JoplinError(xhr.status, xhr.statusText));
        }
      };

      xhr.onerror = () => {
        reject(new JEC_JoplinError(xhr.status, xhr.statusText || 'XmlHttpRequest failed'));
      };

      if (opts.timeout) {
        xhr.timeout = opts.timeout;

        xhr.ontimeout = () => {
          reject(new JEC_JoplinError(504, 'XmlHttpRequest timed out'));
        }
      }

      if (opts.headers) {
        Object.keys(opts.headers).forEach((key) => {
          xhr.setRequestHeader(key, opts.headers[key]);
        });
      }

      let params = opts.params;
      // We'll need to stringify if we've been given an object
      // If we have a string, this is skipped.
      if (params && typeof params === 'object') {
        params = Object.keys(params).map((key) => {
          return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
        }).join('&');
      }

      xhr.send(params);
    });
  }

  async tagNote_(noteId, tagIds) {
    for (const id of tagIds) {
      await this.request_({
        method: 'POST',
        url: 'http://127.0.0.1:' + this.port_.toString() + '/tags/' + id + '/notes',
        headers: { 'Content-Type': 'application/json' },
        params: '{ "id": ' + JSON.stringify(noteId) + ' }',
        timeout: 10000
      });
    }
  }
}

class JEC_Popup {
  constructor() {
    this.window_ = null;
  }

  set body(val) {
    const b = this.window_.document.getElementById('joplin-preview-body');
    b.value = val;
    this.window_.sizeToContent();
  }

  get checkedTagMenuItems_() {
    const list = this.window_.document.getElementById('joplin-tag-list');

    return Array.from(list.childNodes).filter(e => e.hasAttribute('checked'));
  }

  close() {
    this.window_.close();
    this.window_ = null;
  }

  getConfirmation() {
    if (this.window_.closed) {
      return Promise.resolve(false);
    }

    const confirm = new Promise((resolve) => {
      const c = this.window_.document.getElementById('joplin-confirm');
      c.addEventListener('command', function() {
        resolve(true);
      }, { once: true });
      c.disabled = false;
    });

    const cancel = new Promise((resolve) => {
      this.window_.addEventListener('unload', function() {
        resolve(false);
      }, {once: true });
    });

    return Promise.race([confirm, cancel]);
  }

  get notebookId() {
    const menu = this.window_.document.getElementById('joplin-notebook-menu');
    return menu.selectedItem.getAttribute('value');
  }

  set notebooks(val) {
    const menu = this.window_.document.getElementById('joplin-notebook-menu');
    const list = this.window_.document.getElementById('joplin-notebook-list');
    let selection = 0;
    let i = 0;

    val.forEach((e) => {
      const menuItem = document.createElement('menuitem');
      menuItem.setAttribute('label', e.title);
      menuItem.setAttribute('value', e.id);
      list.appendChild(menuItem);

      if (lastSelectedNotebookId && (lastSelectedNotebookId === e.id)) {
        selection = i;
      }

      i++;
    });

    menu.addEventListener('command', () => {
      this.updateSelectedNotebook_();
    });

    menu.selectedIndex = selection;
    menu.disabled = false;
  }

  open() {
    return new Promise((resolve) => {
      this.window_ = window.open(
        'chrome://emailclipper/content/popup.xul',
        'joplin',
        'chrome,resizable,centerscreen,scrollbars');
      this.window_.onload = () => {
        resolve(true);
      };
    })
    .then(() => {
      const c = this.window_.document.getElementById('joplin-cancel');
      c.addEventListener('command', () => {
        this.close();
      }, { once: true });

      return true;
    });
  }

  set status(val) {
    const s = this.window_.document.getElementById('joplin-status');
    s.value = val;
  }

  set subject(val) {
    const s = this.window_.document.getElementById('joplin-preview-subject');
    s.value = val;
  }

  get tagIds() {
    return this.checkedTagMenuItems_.map(e => e.getAttribute('value'));
  }

  set tags(val) {
    const menu = this.window_.document.getElementById('joplin-tag-menu');
    const list = this.window_.document.getElementById('joplin-tag-list');
    const text = this.window_.document.getElementById('joplin-tags');

    val.forEach((e) => {
      const menuItem = document.createElement('menuitem');
      menuItem.setAttribute('label', e.title);
      menuItem.setAttribute('value', e.id);
      menuItem.setAttribute('type', 'checkbox');
      list.appendChild(menuItem);
    });

    menu.addEventListener('command', () => {
      this.updateTagList_();
    });

    menu.disabled = false;
    text.disabled = false;
  }

  updateSelectedNotebook_() {
    lastSelectedNotebookId = this.notebookId;
  }

  updateTagList_() {
    const text = this.window_.document.getElementById('joplin-tags');

    text.value = this.checkedTagMenuItems_.map(e => e.getAttribute('label')).join(', ');
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
    const notebooks = [];

    function helper(item, level) {
      notebooks.push({
        title: ' . '.repeat(level) + item.title,
        id: item.id
      });

      if (item.children) {
        item.children.forEach(e => helper(e, level + 1));
      }
    }

    list.forEach(e => helper(e, 0));

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
    titleBlock.forEach((e) => {
      if (!e.optional || msg[e.prop]) {
        maxLabel = Math.max(maxLabel, e.label.length);
        maxProp = Math.max(maxProp, msg[e.prop].length);
      }
    });

    let note = '| ' + ' '.repeat(maxLabel + 5) + ' | ' + ' '.repeat(maxProp) + ' |\n' +
               '| ' + '-'.repeat(maxLabel + 5) + ' | ' + '-'.repeat(maxProp) + ' |\n';

    titleBlock.forEach((e) => {
      if (!e.optional || msg[e.prop]) {
        note += '| **' + e.label + ':**' + ' '.repeat(maxLabel - e.label.length) +
          ' | ' + msg[e.prop] + ' '.repeat(maxProp - msg[e.prop].length) + ' |\n';
      }
    });

    note += '\n' + msg.plainBody;

    return note;
  }

  async sendToJoplin() {
    await this.popup_.open();

    const msg = this.tbird_.getCurrentMessage();

    this.popup_.subject = msg.subject;
    this.popup_.body = 'Downloading...';
    await msg.download();

    const note = this.messageToNote_(msg);
    this.popup_.body = note;

    await this.connectToJoplin_();
    this.popup_.notebooks = this.flattenNotebookList_(await this.joplin_.getNotebooks());
    this.popup_.tags = await this.joplin_.getTags();

    if (!await this.popup_.getConfirmation()) {
      return false;
    }

    await this.joplin_.createNote(msg.subject, note, this.popup_.notebookId, this.popup_.tagIds);

    this.popup_.close();

    return true;
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
  const clipper = new JEC_EmailClipper();
  clipper.sendToJoplin()
    .then((result) => {
      if (result) {
        console.info('sendToJoplin done');
      }
      else {
        console.info('sendToJoplin cancelled');
      }
    })
    .catch((error) => {
      console.error('sendToJoplin failed: ' + error.toString())
    });
}
