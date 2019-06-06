'use strict';

ChromeUtils.import("resource:///modules/jsmime.jsm");
ChromeUtils.import("resource:///modules/gloda/mimemsg.js");
ChromeUtils.import("resource://gre/modules/FileUtils.jsm");
ChromeUtils.import("resource://gre/modules/osfile.jsm");

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
      { prop: 'date',    label: 'Date',    optional: false },
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
    this.notebooks_ = null;
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

  addSelectedTagToList_() {
    const list = this.window_.document.getElementById('jec-tag-list');
    const item = list.selectedItem;

    if (item) {
      this.addTagToList_(item.getAttribute('label'), item.getAttribute('value'));
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

  populateNotebooksTree_(val, select) {
    const tree = this.window_.document.getElementById('jec-notebooks-tree');
    const list = this.window_.document.getElementById('jec-notebooks-list');

    function helper(notebook, parentElement) {
      notebooks.push(notebook);

      const item = document.createElement('treeitem');
      item.setAttribute('value', notebook.id);
      if (notebook.children) {
        item.setAttribute('container', 'true');
      }

      const row = document.createElement('treerow');

      const cell = document.createElement('treecell');
      cell.setAttribute('label', notebook.title);

      row.appendChild(cell);
      item.appendChild(row);
      parentElement.appendChild(item);

      if (notebook.children) {
        const nextParent = document.createElement('treechildren');
        item.appendChild(nextParent);

        notebook.children.forEach(e => helper(e, nextParent));
      }
    }

    const notebooks = [];
    val.forEach(e => helper(e, list));
    this.notebooks_ = notebooks;

    if (select) {
      tree.view.selection.clearSelection();
      tree.view.selection.select(0);
    }

    return tree;
  }

  populateRecentPicks_(recentIds, select) {
    const recent = this.window_.document.getElementById('jec-notebooks-recent');
    recentIds.forEach((id) => {
      const nb = this.notebooks_.find(n => n.id === id);
      if (nb) {
        const listItem = document.createElement('listitem');
        listItem.setAttribute('label', nb.title);
        listItem.setAttribute('value', nb.id);
        recent.appendChild(listItem);
      }
    });

    recent.setAttribute('rows', recentIds.length);

    if (select) {
      recent.selectedIndex = 0;
    }

    return recent;
  }

  get selectedAttachmentIndices() {
    return this.checkedAttachmentElements_.map(e => e.getAttribute('value'));
  }

  get selectedNotebookId() {
    const recent = this.window_.document.getElementById('jec-notebooks-recent');
    if (recent.selectedItem) {
      return recent.selectedItem.getAttribute('value');
    }

    const tree = this.window_.document.getElementById('jec-notebooks-tree');
    return tree.view.getItemAtIndex(tree.currentIndex).getAttribute('value');
  }

  get selectedTagTitles() {
    const text = this.window_.document.getElementById('jec-tags');

    return Array.from(text.childNodes).map(e => e.getAttribute('label'));
  }

  setAttachments(val) {
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

  setNotebooks(val, recentIds) {
    const selectRecent = recentIds.length > 0;

    const tree = this.populateNotebooksTree_(val, !selectRecent);
    const recent = this.populateRecentPicks_(recentIds, selectRecent);

    tree.addEventListener('select', () => {
      if ((tree.view.selection.count != 0) && (recent.selectedIndex != -1)) {
        recent.clearSelection();
      }
    });

    recent.addEventListener('select', () => {
      if ((recent.selectedIndex != -1) && (tree.view.selection.count != 0)) {
        tree.view.selection.clearSelection();
      }
    });

    recent.disabled = false;
    tree.disabled = false;
  }

  setStatus(val) {
    const s = this.window_.document.getElementById('jec-status');
    s.value = val;
  }

  setTags(val) {
    const list = this.window_.document.getElementById('jec-tag-list');
    const text = this.window_.document.getElementById('jec-tags');

    val.forEach((e) => {
      const listItem = document.createElement('listitem');
      listItem.setAttribute('label', e.title);
      listItem.setAttribute('value', e.id);
      list.appendChild(listItem);
    });

    list.setAttribute('rows', Math.min(val.length, 5));

    list.addEventListener('select', () => {
      this.addSelectedTagToList_();
    });

    text.addEventListener('change', () => {
      this.addManualTagsToList_();
    });

    text.disabled = false;
    list.disabled = false;
  }

  tagIsInList_(title) {
    const text = this.window_.document.getElementById('jec-tags');

    return Array.from(text.childNodes).reduce(
      (t, e) => t || (e.getAttribute('label') === title),
      false);
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

  get date() {
    const d = new Date(this.header_.dateInSeconds * 1000);
    return d.toLocaleString();
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

class JEC_Storage {
  constructor() {
    this.path_ = OS.Path.join(OS.Constants.Path.profileDir, 'joplin-email-clipper-storage.json');
    this.data_ = { recentPicks: [] };
    this.initialized_ = false;
  }

  async init_() {
    if (this.initialized_) {
      return;
    }

    try {
      const data = await OS.File.read(this.path_, { encoding: 'utf-8' });
      this.data_ = Object.assign(this.data_, JSON.parse(data));
    }
    catch (error) {
      if (error.becauseNoSuchFile) {
        // ignore
      }
      else {
        throw error;
      }
    }

    this.initialized_ = true;
  }

  async getRecentPicks() {
    await this.init_();

    return this.data_.recentPicks;
  }

  async setRecentPicks(val) {
    this.data_.recentPicks = val;

    await this.update_();
  }

  async update_() {
    await OS.File.writeAtomic(this.path_, JSON.stringify(this.data_), { encoding: 'utf-8' });
  }
}

class JEC_EmailClipper {
  constructor() {
    this.popup_ = new JEC_Popup();
    this.tbird_ = new JEC_Thunderbird();
    this.joplin_ = new JEC_Joplin();
    this.storage_ = new JEC_Storage();
  }

  async connectToJoplin_() {
    this.popup_.setStatus('Looking for service');

    while (!this.joplin_.connected && !this.popup_.isCancelled) {
      if (!await this.joplin_.connect()) {
        await this.sleep_(1000);
      }
    }

    if (this.joplin_.connected) {
      this.popup_.setStatus('Ready on port ' + this.joplin_.port.toString());
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

  async sendToJoplin() {
    let folder = null;

    await this.popup_.open();

    const msg = this.tbird_.getCurrentMessage();
    await msg.download();

    let attachments = msg.attachments;
    this.popup_.setAttachments(attachments);

    if (!await this.connectToJoplin_()) {
      return false;
    }

    this.popup_.setNotebooks(await this.joplin_.getNotebooks(), await this.storage_.getRecentPicks());
    this.popup_.setTags(await this.joplin_.getTags());

    if (!await this.popup_.getConfirmation()) {
      return false;
    }

    await this.updateRecentPicks_(this.popup_.selectedNotebookId);

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

  async updateRecentPicks_(val) {
    let selections = await this.storage_.getRecentPicks();
    selections = selections.filter(id => id !== val);
    selections.unshift(val);
    await this.storage_.setRecentPicks(selections.slice(0,5));
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
