'use strict';

let EXPORTED_SYMBOLS = ['JEC_EmailClipper'];

let { FileUtils } = ChromeUtils.import('resource://gre/modules/FileUtils.jsm');
let { jsmime } = ChromeUtils.import('resource:///modules/jsmime.jsm');
ChromeUtils.import('resource:///modules/gloda/mimemsg.js');
ChromeUtils.import('resource://gre/modules/osfile.jsm');

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

  addAttachments_(body, attData) {
    if (attData.length !== 0) {
      body += '\n\n';

      attData.forEach((att) => {
        body += '[' + att.attachment.fileName + '](:/' + att.resource.id + ')\n';
      });
    }

    return body;
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

  async createNote(title, body, notebookId, tagTitles, attachments) {
    const attData = [];
    for (const att of attachments) {
      attData.push({
        attachment: att,
        resource: await this.createResource_(await att.asFile(), 'Email attachment', att.fileName)
      });
    }

    const bodyWithAttachments = this.addAttachments_(body, attData);

    await this.request_({
      method: 'POST',
      url: this.url_('notes'),
      headers: { 'Content-Type': 'application/json' },
      params: '{ "title": ' + JSON.stringify(title) +
              ', "body": ' + JSON.stringify(bodyWithAttachments) +
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

class JEC_TagsPicker {
  constructor(win) {
    this.text_ = win.document.getElementById('jec-tags');
    this.list_ = win.document.getElementById('jec-tag-list');
  }

  addManualTagsToList_() {
    this.text_.value.split(',').forEach((t) => {
      const title = t.trim();
      if (title) {
        this.addTagToList_(title, '');
      }
    });

    this.text_.value = '';
  }

  addSelectedTagToList_() {
    const item = this.list_.selectedItem;

    if (item) {
      this.addTagToList_(item.getAttribute('label'), item.getAttribute('value'));
      this.list_.clearSelection();
    }
  }

  addTagToList_(title, id) {
    if (!this.tagIsInList_(title)) {
      const button = document.createElement('button');
      button.setAttribute('label', title);
      button.setAttribute('value', id);
      button.setAttribute('image', 'chrome://emailclipper/content/red_x_icon.png');

      button.addEventListener('command', function () {
        // 'this' is bound to button element
        this.parentElement.removeChild(this);
      });

      this.text_.appendChild(button);
    }
  }

  get selectedTagTitles() {
    return Array.from(this.text_.childNodes).map(e => e.getAttribute('label'));
  }

  setTags(val) {
    val.forEach((e) => {
      const listItem = document.createElement('listitem');
      listItem.setAttribute('label', e.title);
      listItem.setAttribute('value', e.id);
      this.list_.appendChild(listItem);
    });

    this.list_.setAttribute('rows', Math.min(val.length, 5));

    this.list_.addEventListener('select', () => {
      this.addSelectedTagToList_();
    });

    this.text_.addEventListener('change', () => {
      this.addManualTagsToList_();
    });

    this.text_.disabled = false;
    this.list_.disabled = false;
  }

  tagIsInList_(title) {
    return Array.from(this.text_.childNodes).reduce(
      (t, e) => t || (e.getAttribute('label') === title),
      false);
  }
}

class JEC_NotebookPicker {
  constructor(win, onGotSelection, onLostSelection) {
    this.deck_ = win.document.getElementById('jec-notebooks-deck');
    this.tree_ = win.document.getElementById('jec-notebooks-tree');
    this.list_ = win.document.getElementById('jec-notebooks-list');
    this.recent_ = win.document.getElementById('jec-notebooks-recent');
    this.search_ = win.document.getElementById('jec-notebooks-search');
    this.results_ = win.document.getElementById('jec-notebooks-results');
    this.onGotSelection_ = onGotSelection;
    this.onLostSelection_ = onLostSelection;
    this.notebooks_ = null;
  }

  hasSelection() {
    return this.selectedNotebookId !== null;
  }

  populateNotebooksTree_(val) {
    const notebooks = [];

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

    val.forEach(e => helper(e, this.list_));
    this.notebooks_ = notebooks;
  }

  populateRecentPicks_(recentIds) {
    recentIds.forEach((id) => {
      const nb = this.notebooks_.find(n => n.id === id);
      if (nb) {
        const listItem = document.createElement('listitem');
        listItem.setAttribute('label', nb.title);
        listItem.setAttribute('value', nb.id);
        this.recent_.appendChild(listItem);
      }
    });

    this.recent_.setAttribute('rows', recentIds.length);
  }

  get selectedNotebookId() {
    // deck.selectedIndex returns a string, not an integer
    if (Number(this.deck_.selectedIndex) === 0) {
      const selectedItem = this.recent_.selectedItem;
      if (selectedItem) {
        return selectedItem.getAttribute('value');
      }
      else {
        return this.tree_.view.getItemAtIndex(this.tree_.currentIndex).getAttribute('value');
      }
    }
    else {
      const selectedItem = this.results_.selectedItem;
      if (selectedItem) {
        return selectedItem.getAttribute('value');
      }
      else {
        return null;
      }
    }
  }

  setNotebooks(val, recentIds) {
    this.populateNotebooksTree_(val);
    this.populateRecentPicks_(recentIds);

    if (this.recent_.hasChildNodes()) {
      this.recent_.selectedIndex = 0;
    }
    else {
      this.tree_.view.selection.clearSelection();
      this.tree_.view.selection.select(0);
    }

    this.tree_.addEventListener('select', () => {
      if ((this.tree_.view.selection.count !== 0) && (this.recent_.selectedIndex !== -1)) {
        this.recent_.clearSelection();
      }
    });

    this.recent_.addEventListener('select', () => {
      if ((this.recent_.selectedIndex !== -1) && (this.tree_.view.selection.count !== 0)) {
        this.tree_.view.selection.clearSelection();
      }
    });

    this.search_.addEventListener('command', () => {
      this.updateNotebookSearchResults_();
    });

    this.results_.addEventListener('select', () => {
      if (this.results_.selectedIndex !== -1) {
        this.onGotSelection_();
      }
      else {
        this.onLostSelection_();
      }
    });

    this.recent_.disabled = false;
    this.tree_.disabled = false;
    this.search_.disabled = false;
  }

  updateNotebookSearchResults_() {
    let query = this.search_.value;
    if (!query) {
      this.results_.disabled = true;
      this.deck_.selectedIndex = 0;
      this.onGotSelection_();
      return;
    }

    query = query.toLocaleLowerCase();
    this.deck_.selectedIndex = 1;
    this.onLostSelection_();

    Array.from(this.results_.childNodes).forEach((c) => {
      this.results_.removeChild(c);
    });

    const answers = this.notebooks_.filter(n => n.title.toLocaleLowerCase().includes(query));
    if (answers.length > 0) {
      answers.forEach((nb) => {
        const listItem = document.createElement('listitem');
        listItem.setAttribute('label', nb.title);
        listItem.setAttribute('value', nb.id);
        this.results_.appendChild(listItem);
      });

      this.results_.disabled = false;
    }
    else {
      const listItem = document.createElement('listitem');
      listItem.setAttribute('label', 'No matches found');
      this.results_.appendChild(listItem);
      this.results_.disabled = true;
    }
  }
}

class JEC_AttachmentsPicker {
  constructor(win) {
    this.list_ = win.document.getElementById('jec-attachments');
  }

  get checkedAttachmentElements_() {
    return Array.from(this.list_.childNodes).filter(e => e.hasAttribute('checked'));
  }

  get selectedAttachmentIndices() {
    return this.checkedAttachmentElements_.map(e => e.getAttribute('value'));
  }

  setAttachments(val) {
    if (val.length !== 0) {
      for (let i = 0; i < val.length; i++) {
        const checkBox = document.createElement('checkbox');
        checkBox.setAttribute('label', val[i].fileName);
        checkBox.setAttribute('value', i.toString());
        checkBox.setAttribute('checked', 'true');
        this.list_.appendChild(checkBox);
      }
    }
    else {
      const label = document.createElement('label');
      label.setAttribute('value', 'None');
      this.list_.appendChild(label);
    }
  }
}

class JEC_OptionsPicker {
  constructor(win) {
    this.optionFormat_ = win.document.getElementById('jec-option-format-title');
    this.titleFormat_ = win.document.getElementById('jec-title-format');
  }

  getTitleFormat_() {
    if (this.optionFormat_.checked) {
      return this.titleFormat_.value;
    }
    else {
      return null;
    }
  }

  get options() {
    return { titleFormat: this.getTitleFormat_() };
  }

  set options(val) {
    this.setTitleFormat_(val.titleFormat);
  }

  setTitleFormat_(val) {
    if (val) {
      this.titleFormat_.value = val;
      this.titleFormat_.disabled = false;
      this.optionFormat_.checked = true;
    }
    else {
      this.titleFormat_.disabled = true;
      this.optionFormat_.checked = false;
    }

    this.optionFormat_.addEventListener('command', () => {
      if (this.optionFormat_.checked) {
        this.titleFormat_.disabled = false;
      }
      else {
        this.titleFormat_.disabled = true;
      }
    });
  }
}

class JEC_StatusBar {
  constructor(win) {
    this.status_ = win.document.getElementById('jec-status');
    this.joplinStatus_ = '';
    this.tbirdStatus_ = '';
    this.userStatus_ = '';
  }

  set status(val) {
    if ('joplin' in val) {
      this.joplinStatus_ = val.joplin;
    }

    if ('tbird' in val) {
      this.tbirdStatus_ = val.tbird;
    }

    if ('user' in val) {
      this.userStatus_ = val.user;
    }

    this.status_.value = [this.userStatus_, this.tbirdStatus_, this.joplinStatus_].join(' ');
  }
}

class JEC_Popup {
  constructor() {
    this.window_ = null;
    this.notebook_ = null;
    this.tags_ = null;
    this.attachments_ = null;
    this.options_ = null;
    this.confirm_ = null;
    this.cancel_ = null;
    this.status_ = null;
    this.waitingForConfirmation_ = false;
  }

  close() {
    this.window_.close();
    this.window_ = null;
  }

  async getConfirmation() {
    if (this.isCancelled()) {
      return Promise.resolve(false);
    }

    this.waitingForConfirmation_ = true;

    const confirm = new Promise((resolve) => {
      this.confirm_.addEventListener('command', function() {
        resolve(true);
      }, { once: true });

      this.confirm_.disabled = !this.notebook_.hasSelection();
    });

    const cancel = new Promise((resolve) => {
      this.window_.addEventListener('unload', function() {
        resolve(false);
      }, { once: true });
    });

    return Promise.race([confirm, cancel]);
  }

  isCancelled() {
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
      this.confirm_ = this.window_.document.getElementById('jec-confirm');
      this.cancel_ = this.window_.document.getElementById('jec-cancel');
      this.notebook_ = new JEC_NotebookPicker(this.window_,
        // onGotSelection
        () => {
          this.confirm_.disabled = !this.waitingForConfirmation_;
        },
        // onLostSelection
        () => {
          this.confirm_.disabled = true;
        });
      this.tags_ = new JEC_TagsPicker(this.window_);
      this.attachments_ = new JEC_AttachmentsPicker(this.window_);
      this.options_ = new JEC_OptionsPicker(this.window_);
      this.status_ = new JEC_StatusBar(this.window_);

      this.cancel_.addEventListener('command', () => {
        this.close();
      }, { once: true });

      return true;
    });
  }

  get options() {
    return this.options_.options;
  }

  get selectedAttachmentIndices() {
    return this.attachments_.selectedAttachmentIndices;
  }

  get selectedNotebookId() {
    return this.notebook_.selectedNotebookId;
  }

  get selectedTagTitles() {
    return this.tags_.selectedTagTitles;
  }

  setAttachments(val) {
    this.attachments_.setAttachments(val);
  }

  setNotebooks(val, recentIds) {
    this.notebook_.setNotebooks(val, recentIds);
  }

  setOptions(val) {
    this.options_.options = val;
  }

  setStatus(val) {
    this.status_.status = val;
  }

  setTags(val) {
    this.tags_.setTags(val);
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
      this.file_.createUnique(Ci.nsIFile.NORMAL_FILE_TYPE, FileUtils.PERMS_FILE);

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

  get localeDateTime() {
    const d = new Date(this.header_.dateInSeconds * 1000);
    return d.toLocaleString();
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

  get ymdDateTime() {
    function pad(number) {
      if (number < 10) {
        return '0' + number.toString();
      }

      return number.toString();
    }

    const d = new Date(this.header_.dateInSeconds * 1000);
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' +
              pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  }
}

class JEC_Thunderbird {
  constructor() {
  }

  createTemporaryFolder(name) {
    const folder = FileUtils.getFile("TmpD", [name]);
    folder.createUnique(Ci.nsIFile.DIRECTORY_TYPE, FileUtils.PERMS_DIRECTORY);
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
  constructor(defaults) {
    this.path_ = OS.Path.join(OS.Constants.Path.profileDir, 'joplin-email-clipper-storage.json');
    this.data_ = defaults;
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

  async getOptions() {
    await this.init_();

    return this.data_.options;
  }

  async getRecentPicks() {
    await this.init_();

    return this.data_.recentPicks;
  }

  async setOptions(val) {
    this.data_.options = val;

    await this.update_();
  }

  async setRecentPicks(val) {
    this.data_.recentPicks = val;

    await this.update_();
  }

  async update_() {
    await OS.File.writeAtomic(this.path_, JSON.stringify(this.data_), { encoding: 'utf-8' });
  }
}

class JEC_Formatter {
  static format(input, data, whiteList) {
    const FIELDSPEC_REGEXP = /(\${[a-zA-Z]+?})/;
    const SPLIT_REGEXP = /\${([a-zA-Z]+?)}/;

    const parts = input.split(FIELDSPEC_REGEXP);

    return parts.reduce((acc, p) => {
      const field = SPLIT_REGEXP.exec(p);
      if (field && (whiteList.indexOf(field[1]) != -1)) {
        return acc + data[field[1]];
      }
      else {
        return acc + p;
      }
    }, '');
  }
}

class JEC_EmailClipper {
  constructor() {
    this.popup_ = new JEC_Popup();
    this.tbird_ = new JEC_Thunderbird();
    this.joplin_ = new JEC_Joplin();
    this.storage_ = new JEC_Storage({ recentPicks: [],
                                      options: { titleFormat: '' } });
  }

  async connectToJoplin_() {
    while (!this.joplin_.connected && !this.popup_.isCancelled()) {
      if (!await this.joplin_.connect()) {
        await this.sleep_(1000);
      }
    }

    if (this.joplin_.connected) {
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

  formatMessage_(message) {
    const titleBlock = [
      { prop: 'localeDateTime', label: 'Date',    optional: false },
      { prop: 'from',           label: 'From',    optional: false },
      { prop: 'subject',        label: 'Subject', optional: false },
      { prop: 'to',             label: 'To',      optional: false },
      { prop: 'cc',             label: 'Cc',      optional: true  },
      { prop: 'bcc',            label: 'Bcc',     optional: true  }
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

    return body;
  }

  async sendToJoplin() {
    let folder = null;

    await this.popup_.open();

    this.popup_.setStatus({ tbird: 'Downloading message.' });

    const msg = this.tbird_.getCurrentMessage();
    await msg.download();

    let attachments = msg.attachments;
    this.popup_.setAttachments(attachments);

    this.popup_.setStatus({ tbird: '', joplin: 'Looking for Joplin service.' });

    if (!await this.connectToJoplin_()) {
      return false;
    }

    this.popup_.setStatus({ joplin: 'Found Joplin on port ' + this.joplin_.port.toString() +
                                    ', getting notebooks and tags.' });

    this.popup_.setNotebooks(await this.joplin_.getNotebooks(), await this.storage_.getRecentPicks());
    this.popup_.setTags(await this.joplin_.getTags());
    this.popup_.setOptions(await this.storage_.getOptions());

    this.popup_.setStatus({ joplin: '', user: 'Waiting for confirmation.' });

    if (!await this.popup_.getConfirmation()) {
      return false;
    }

    this.popup_.setStatus({ user: '' });

    await this.updateRecentPicks_(this.popup_.selectedNotebookId);
    const options = this.popup_.options;
    await this.storage_.setOptions(options);

    let selectedAttachments = this.popup_.selectedAttachmentIndices.map(i => attachments[i]);
    if (selectedAttachments.length !== 0) {
      this.popup_.setStatus({ tbird: 'Downloading attachment(s).' });

      folder = this.tbird_.createTemporaryFolder('jec-attachments');
      await this.downloadAttachments_(selectedAttachments, folder);

      this.popup_.setStatus({ tbird: '' });
    }

    this.popup_.setStatus({ joplin: 'Creating note.' });

    let title = msg.subject;
    if (options.titleFormat) {
      title = JEC_Formatter.format(options.titleFormat, msg,
                ['bcc', 'cc', 'from', 'localeDateTime', 'subject', 'to', 'ymdDateTime']);
    }

    const body = this.formatMessage_(msg);

    await this.joplin_.createNote(
      title,
      body,
      this.popup_.selectedNotebookId,
      this.popup_.selectedTagTitles,
      selectedAttachments);

    this.popup_.setStatus({ joplin: 'Note created successfully.' });

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
