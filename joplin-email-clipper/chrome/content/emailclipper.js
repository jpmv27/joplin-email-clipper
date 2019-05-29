'use strict';

ChromeUtils.import("resource://gre/modules/FileUtils.jsm");
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

    for (this.port_ = MINPORT; this.port_ <= MAXPORT; this.port_++) {
      try {
        const response = await this.request_({
          method: 'GET',
          url: this.url_('ping'),
          timeout: 10000
        });

        if (response === 'JoplinClipperServer') {
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

  async createNote(message, notebookId, tagTitles, attachments) {
    const attData = [];
    for (const att of attachments) {
      attData.push({
        attachment: att,
        resource: await this.createResource_(await att.asFile(), 'Email attachment', att.fileName)
      });
    }

    const body = this.formatBody_(message, attData);

    await this.request_({
      method: 'POST',
      url: this.url_('notes'),
      headers: { 'Content-Type': 'application/json' },
      params: '{ "title": ' + JSON.stringify(message.subject) +
              ', "body": ' + JSON.stringify(body) +
              ', "parent_id": ' + JSON.stringify(notebookId) +
              ', "tags": ' + JSON.stringify(tagTitles.join(',')) + ' }',
      timeout: 10000
    });
  }

  async createResource_(data, title, fileName) {
    const formData = new FormData();
    formData.append("data", data);
    formData.append("props", '{ "title": ' + JSON.stringify(title) +
                             ', "filename": ' + JSON.stringify(fileName) + ' }');

    const response = await this.request_({
      method: 'POST',
      url: this.url_('resources'),
      params: formData,
      timeout: 10000
    });

    return JSON.parse(response);
  }

  get connected() {
    return (this.port_ !== -1);
  }

  formatBody_(message, attData) {
    const titleBlock = [
      { prop: 'from',    label: 'From',    optional: false },
      { prop: 'subject', label: 'Subject', optional: false },
      { prop: 'to',      label: 'To',      optional: false },
      { prop: 'cc',      label: 'Cc',      optional: true  },
      { prop: 'bcc',     label: 'Bcc',     optional: true  }
    ];

    // Find maximum length of labels and properties
    let maxLabel = 0;
    let maxProp = 0;
    titleBlock.forEach((e) => {
      if (!e.optional || message[e.prop]) {
        maxLabel = Math.max(maxLabel, e.label.length);
        maxProp = Math.max(maxProp, message[e.prop].length);
      }
    });

    let body = '| ' + ' '.repeat(maxLabel + 5) + ' | ' + ' '.repeat(maxProp) + ' |\n' +
               '| ' + '-'.repeat(maxLabel + 5) + ' | ' + '-'.repeat(maxProp) + ' |\n';

    titleBlock.forEach((e) => {
      if (!e.optional || message[e.prop]) {
        body += '| **' + e.label + ':**' + ' '.repeat(maxLabel - e.label.length) +
          ' | ' + message[e.prop] + ' '.repeat(maxProp - message[e.prop].length) + ' |\n';
      }
    });

    body += '\n' + message.plainBody;

    if (attData.length !== 0) {
      body += '\n\n';

      attData.forEach((att) => {
        body += '[' + att.attachment.fileName + '](:/' + att.resource.id + ')\n';
      });
    }

    return body;
  }

  async getNotebooks() {
    const response = await this.request_({
      method: 'GET',
      url: this.url_('folders'),
      timeout: 10000
    });

    return JSON.parse(response);
  }

  async getTags() {
    const response = await this.request_({
      method: 'GET',
      url: this.url_('tags'),
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

  async request_(opts) {
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

      xhr.send(opts.params);
    });
  }

  url_(endpoint) {
    return 'http://127.0.0.1:' + this.port_.toString() + '/' + endpoint;
  }
}

class JEC_Popup {
  constructor() {
    this.window_ = null;
  }

  addManualTagsToList_() {
    const text = this.window_.document.getElementById('jec-tags');

    text.value.split(',').forEach((t) => {
      const title = t.trim();
      if (title) {
        this.addTagToList_(title, '');
      }
    });

    text.value = '';
  }

  addSelectedTagsToList_() {
    const list = this.window_.document.getElementById('jec-tag-list');
    let selected = false;

    Array.from(list.childNodes).filter(e => e.getAttribute('selected')).forEach((e) => {
      this.addTagToList_(e.getAttribute('label'), e.getAttribute('value'));
      selected = true;
    });

    // clearSelection generates a 'selected' event
    // so we need to avoid infinite recursion
    if (selected) {
      list.clearSelection();
    }
  }

  addTagToList_(title, id) {
    const text = this.window_.document.getElementById('jec-tags');

    if (!this.tagIsInList_(title)) {
      const button = document.createElement('button');
      button.setAttribute('label', title);
      button.setAttribute('value', id);
      button.setAttribute('image', 'chrome://emailclipper/content/red_x_icon.png');

      button.addEventListener('command', function () {
        // 'this' is bound to button element
        this.parentElement.removeChild(this);
      });

      text.appendChild(button);
    }
  }

  set attachments(val) {
    const a = this.window_.document.getElementById('jec-attachments');
    if (val.length !== 0) {
      for (let i = 0; i < val.length; i++) {
        const checkBox = document.createElement('checkbox');
        checkBox.setAttribute('label', val[i].fileName);
        checkBox.setAttribute('value', i.toString());
        checkBox.setAttribute('checked', 'true');
        a.appendChild(checkBox);
      }
    }
    else {
      const label = document.createElement('label');
      label.setAttribute('value', 'None');
      a.appendChild(label);
    }
  }

  get checkedAttachmentElements_() {
    const list = this.window_.document.getElementById('jec-attachments');

    return Array.from(list.childNodes).filter(e => e.hasAttribute('checked'));
  }

  close() {
    this.window_.close();
    this.window_ = null;
  }

  getConfirmation() {
    if (this.isCancelled) {
      return Promise.resolve(false);
    }

    const confirm = new Promise((resolve) => {
      const c = this.window_.document.getElementById('jec-confirm');
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

  get isCancelled() {
    return !this.window_ || this.window_.closed;
  }

  set notebooks(val) {
    const menu = this.window_.document.getElementById('jec-notebook-menu');
    const list = this.window_.document.getElementById('jec-notebook-list');
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

  async open() {
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
      const c = this.window_.document.getElementById('jec-cancel');
      c.addEventListener('command', () => {
        this.close();
      }, { once: true });

      return true;
    });
  }

  get selectedAttachmentIndices() {
    return this.checkedAttachmentElements_.map(e => e.getAttribute('value'));
  }

  get selectedNotebookId() {
    const menu = this.window_.document.getElementById('jec-notebook-menu');
    return menu.selectedItem.getAttribute('value');
  }

  get selectedTagTitles() {
    const text = this.window_.document.getElementById('jec-tags');

    return Array.from(text.childNodes).map(e => e.getAttribute('label'));
  }

  set status(val) {
    const s = this.window_.document.getElementById('jec-status');
    s.value = val;
  }

  tagIsInList_(title) {
    const text = this.window_.document.getElementById('jec-tags');

    return Array.from(text.childNodes).reduce(
      (t, e) => t || (e.getAttribute('label') === title),
      false);
  }

  set tags(val) {
    const list = this.window_.document.getElementById('jec-tag-list');
    const text = this.window_.document.getElementById('jec-tags');

    val.forEach((e) => {
      const listItem = document.createElement('listitem');
      listItem.setAttribute('label', e.title);
      listItem.setAttribute('value', e.id);
      list.appendChild(listItem);
    });

    list.addEventListener('select', () => {
      this.addSelectedTagsToList_();
    });

    text.addEventListener('change', () => {
      this.addManualTagsToList_();
    });

    text.disabled = false;
    list.disabled = false;
  }

  updateSelectedNotebook_() {
    lastSelectedNotebookId = this.selectedNotebookId;
  }
}

class JEC_Attachment {
  constructor(attachment, header) {
    this.attachment_ = attachment;
    this.header_ = header;
    this.file_ = null;
  }

  async asFile() {
    return await File.createFromNsIFile(this.file_);
  }

  async downloadToFolder(folder) {
    return new Promise((resolve) => {
      this.file_ = folder.clone();
      this.file_.append(this.attachment_.name);
      this.file_.createUnique(Components.interfaces.nsIFile.NORMAL_FILE_TYPE, FileUtils.PERMS_FILE);

      messenger.saveAttachmentToFile(
        this.file_,
        this.attachment_.url,
        this.header_.folder.getUriForMsg(this.header_),
        this.attachment_.contentType,
        {
          OnStartRunningUrl: () => {},
          OnStopRunningUrl: () => {
            resolve(true);
          }
        });
    });
  }

  get fileName() {
    return this.attachment_.name;
  }
}

class JEC_Message {
  constructor(header) {
    this.header_ = header;
    this.message_ = null;
  }

  get attachments() {
    return this.message_.allUserAttachments.map(att => new JEC_Attachment(att, this.header_));
  }

  get bcc() {
    return jsmime.headerparser.decodeRFC2047Words(this.header_.bccList);
  }

  get cc() {
    return jsmime.headerparser.decodeRFC2047Words(this.header_.ccList);
  }

  async download() {
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

  createTemporaryFolder(name) {
    const folder = FileUtils.getFile("TmpD", [name]);
    folder.createUnique(Components.interfaces.nsIFile.DIRECTORY_TYPE, FileUtils.PERMS_DIRECTORY);
    return folder;
  }

  deleteTemporaryFolder(folder) {
    folder.remove(true /* recursive */);
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
    this.popup_.status = 'Looking for service';

    while (!this.joplin_.connected && !this.popup_.isCancelled) {
      if (!await this.joplin_.connect()) {
        await this.sleep_(1000);
      }
    }

    if (this.joplin_.connected) {
      this.popup_.status = 'Ready on port ' + this.joplin_.port.toString();
      return true;
    }
    else {
      return false;
    }
  }

  async downloadAttachments_(attachments, folder) {
    for (const att of attachments) {
      await att.downloadToFolder(folder);
    }
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

  async sendToJoplin() {
    let folder = null;

    await this.popup_.open();

    const msg = this.tbird_.getCurrentMessage();
    await msg.download();

    let attachments = msg.attachments;
    this.popup_.attachments = attachments;

    if (!await this.connectToJoplin_()) {
      return false;
    }

    this.popup_.notebooks = this.flattenNotebookList_(await this.joplin_.getNotebooks());
    this.popup_.tags = await this.joplin_.getTags();

    if (!await this.popup_.getConfirmation()) {
      return false;
    }

    let selectedAttachments = this.popup_.selectedAttachmentIndices.map(i => attachments[i]);
    if (selectedAttachments.length !== 0) {
      folder = this.tbird_.createTemporaryFolder('jec-attachments');
      await this.downloadAttachments_(selectedAttachments, folder);
    }

    await this.joplin_.createNote(
      msg,
      this.popup_.selectedNotebookId,
      this.popup_.selectedTagTitles,
      selectedAttachments);

    attachments = null;
    if (selectedAttachments.length !== 0) {
      selectedAttachments = null;
      this.tbird_.deleteTemporaryFolder(folder);
      folder = null;
    }

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
