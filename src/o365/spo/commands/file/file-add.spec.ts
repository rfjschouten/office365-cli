import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./file-add');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
// import * as fs from 'fs';

describe(commands.FILE_ADD, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  const expectedTitle = `List Item 1`;

  const expectedId = 147;
  let actualId = 0;

  const expectedContentType = 'Item';
  let actualContentType = '';

  let postFakes = (opts: any) => {
    if (opts.url.indexOf('/common/oauth2/token') > -1) {
      return Promise.resolve('abc');
    }

    if (opts.url.indexOf('/_api/contextinfo') > -1) {
      return Promise.resolve({
        FormDigestValue: 'abc'
      });
    }
    if (opts.url.indexOf('ValidateUpdateListItem') > -1) {

      const bodyString = JSON.stringify(opts.body);
      const ctMatch = bodyString.match(/\"?FieldName\"?:\s*\"?ContentType\"?,\s*\"?FieldValue\"?:\s*\"?(\w*)\"?/i);
      actualContentType = ctMatch ? ctMatch[1] : "";
      if (bodyString.indexOf("fail updating me") > -1) return Promise.resolve({ value: [{ErrorMessage: 'failed updating'}] })
      return Promise.resolve({ value: [ { ItemId: expectedId }] });

    }
    return Promise.reject('Invalid request');
  }

  let getFakes = (opts: any) => {
    if (opts.url.indexOf('contenttypes') > -1) {
      return Promise.resolve({ value: [ {Id: { StringValue: expectedContentType }, Name: "Item" } ] });
    }
    if (opts.url.indexOf('/items(') > -1) {
      actualId = opts.url.match(/\/items\((\d+)\)/i)[1];
      return Promise.resolve(
        {
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-03-15T10:43:10Z",
          "EditorId": 3,
          "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
          "ID": actualId,
          "Modified": "2018-03-15T10:52:10Z",
          "Title": expectedTitle,
        }
      );
    }
    return Promise.reject('Invalid request');
  }
  // let stubAuth: any = () => {
  //   sinon.stub(request, 'post').callsFake((opts) => {
  //     if (opts.url.indexOf('/common/oauth2/token') > -1) {
  //       return Promise.resolve('abc');
  //     }

  //     return Promise.reject('Invalid request');
  //   });
  // }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });

  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.FILE_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.FILE_ADD);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  // it('command correctly handles file add reject request', (done) => {
  //   sinon.stub(request, 'post').callsFake((opts) => {
  //     if (opts.url.indexOf('/_api/contextinfo') > -1) {
  //       return Promise.resolve({
  //         FormDigestValue: 'abc'
  //       });
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   const err = 'Invalid request';
  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
  //       return Promise.reject(err);
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso.sharepoint.com';
  //   cmdInstance.action = command.action();

  //   cmdInstance.action({
  //     options: {
  //       debug: true,
  //       webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //       folder: '',
  //       path: '/Users/user/documents'
  //     }
  //   }, () => {

  //     try {
  //       assert(cmdInstanceLogSpy.calledWith(new CommandError(err)));
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore([
  //         request.post,
  //         request.get
  //       ]);
  //     }
  //   });
  // });

  // it('uses correct API url when output json option is passed', (done) => {
  //   stubAuth();

  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('select123=') > -1) {
  //       return Promise.resolve('Correct Url1')
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso.sharepoint.com';
  //   cmdInstance.action = command.action();

  //   cmdInstance.action({
  //     options: {
  //       output: 'json',
  //       debug: false,
  //       webUrl: 'https://contoso.sharepoint.com',
  //       id: 'b2307a39-e878-458b-bc90-03bc578531d6'
  //     }
  //   }, () => {

  //     try {
  //       assert('Correct Url');
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore([
  //         request.post,
  //         request.get
  //       ]);
  //     }
  //   });

  // });

  // it('retrieves file as binary string object', (done) => {
  //   stubAuth();
  //   let returnValue: string = 'BinaryFileString';
  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('/_api/web/GetFileById(') > -1) {
  //       return Promise.resolve(returnValue);
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso-admin.sharepoint.com';
  //   auth.site.tenantId = 'abc';
  //   cmdInstance.action = command.action();
  //   cmdInstance.action({
  //     options: {
  //       debug: false,
  //       id: 'b2307a39-e878-458b-bc90-03bc578531d6',
  //       webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //       asString: true
  //     }
  //   }, () => {
  //     try {
  //       assert(cmdInstanceLogSpy.calledWith(returnValue));
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore(request.get);
  //       Utils.restore(request.post);
  //     }
  //   });
  // });

  // it('retrieves and prints all details of file as ListItem object', (done) => {
  //   stubAuth();

  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('?$expand=ListItemAllFields') > -1) {
  //       return Promise.resolve({
  //         "ListItemAllFields": {
  //           "FileSystemObjectType": 0,
  //           "Id": 4,
  //           "ServerRedirectedEmbedUri": "https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview",
  //           "ServerRedirectedEmbedUrl": "https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview",
  //           "ContentTypeId": "0x0101008E462E3ACE8DB844B3BEBF9473311889",
  //           "ComplianceAssetId": null,
  //           "Title": null,
  //           "ID": 4,
  //           "Created": "2018-02-05T09:42:36",
  //           "AuthorId": 1,
  //           "Modified": "2018-02-05T09:44:03",
  //           "EditorId": 1,
  //           "OData__CopySource": null,
  //           "CheckoutUserId": null,
  //           "OData__UIVersionString": "3.0",
  //           "GUID": "2054f49e-0f76-46d4-ac55-50e1c057941c"
  //         },
  //         "CheckInComment": "",
  //         "CheckOutType": 2,
  //         "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
  //         "CustomizedPageStatus": 0,
  //         "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
  //         "Exists": true,
  //         "IrmEnabled": false,
  //         "Length": "331673",
  //         "Level": 1,
  //         "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b",
  //         "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b",
  //         "MajorVersion": 3,
  //         "MinorVersion": 0,
  //         "Name": "Opendag maart 2018.docx",
  //         "ServerRelativeUrl": "/sites/project-x/Documents/Test1.docx",
  //         "TimeCreated": "2018-02-05T08:42:36Z",
  //         "TimeLastModified": "2018-02-05T08:44:03Z",
  //         "Title": "",
  //         "UIVersion": 1536,
  //         "UIVersionLabel": "3.0",
  //         "UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"
  //       });
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso-admin.sharepoint.com';
  //   auth.site.tenantId = 'abc';
  //   cmdInstance.action = command.action();
  //   cmdInstance.action({
  //     options: {
  //       debug: true,
  //       id: 'b2307a39-e878-458b-bc90-03bc578531d6',
  //       webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //       asListItem: true
  //     }
  //   }, () => {
  //     try {
  //       assert(cmdInstanceLogSpy.calledWith({
  //         "ListItemAllFields": {
  //           "FileSystemObjectType": 0,
  //           "Id": 4,
  //           "ServerRedirectedEmbedUri": "https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview",
  //           "ServerRedirectedEmbedUrl": "https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview",
  //           "ContentTypeId": "0x0101008E462E3ACE8DB844B3BEBF9473311889",
  //           "ComplianceAssetId": null,
  //           "Title": null,
  //           "ID": 4,
  //           "Created": "2018-02-05T09:42:36",
  //           "AuthorId": 1,
  //           "Modified": "2018-02-05T09:44:03",
  //           "EditorId": 1,
  //           "OData__CopySource": null,
  //           "CheckoutUserId": null,
  //           "OData__UIVersionString": "3.0",
  //           "GUID": "2054f49e-0f76-46d4-ac55-50e1c057941c"
  //         },
  //         "CheckInComment": "",
  //         "CheckOutType": 2,
  //         "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
  //         "CustomizedPageStatus": 0,
  //         "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
  //         "Exists": true,
  //         "IrmEnabled": false,
  //         "Length": "331673",
  //         "Level": 1,
  //         "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b",
  //         "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b",
  //         "MajorVersion": 3,
  //         "MinorVersion": 0,
  //         "Name": "Opendag maart 2018.docx",
  //         "ServerRelativeUrl": "/sites/project-x/Documents/Test1.docx",
  //         "TimeCreated": "2018-02-05T08:42:36Z",
  //         "TimeLastModified": "2018-02-05T08:44:03Z",
  //         "Title": "",
  //         "UIVersion": 1536,
  //         "UIVersionLabel": "3.0",
  //         "UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"
  //       }));
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore(request.get);
  //       Utils.restore(request.post);
  //     }
  //   });
  // });

  // it('uses correct API url when id option is passed', (done) => {
  //   stubAuth();

  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('/_api/web/GetFileById(') > -1) {
  //       return Promise.resolve('Correct Url')
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso.sharepoint.com';
  //   cmdInstance.action = command.action();

  //   const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

  //   cmdInstance.action({
  //     options: {
  //       debug: false,
  //       id: actionId,
  //       webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //     }
  //   }, () => {

  //     try {
  //       assert(1 === 1);
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore([
  //         request.post,
  //         request.get
  //       ]);
  //     }
  //   });
  // });

  // it('uses correct API url when url option is passed', (done) => {
  //   stubAuth();

  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('/_api/web/GetFileByServerRelativeUrl(') > -1) {
  //       return Promise.resolve('Correct Url')
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso.sharepoint.com';
  //   cmdInstance.action = command.action();

  //   cmdInstance.action({
  //     options: {
  //       debug: false,
  //       url: '/sites/project-x/Documents/Test1.docx',
  //       webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //     }
  //   }, () => {

  //     try {
  //       assert(1 === 1);
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore([
  //         request.post,
  //         request.get
  //       ]);
  //     }
  //   });
  // });

  // it('uses correct API url when url and id are both not passed', (done) => {
  //   stubAuth();

  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url === '') {
  //       return Promise.resolve('Correct Url')
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso.sharepoint.com';
  //   cmdInstance.action = command.action();

  //   cmdInstance.action({
  //     options: {
  //       debug: false,
  //       webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //     }
  //   }, () => {

  //     try {
  //       assert(1 === 1);
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore([
  //         request.post,
  //         request.get
  //       ]);
  //     }
  //   });
  // });

  // it('fails validation if path doesn\'t exist', () => {
  //   sinon.stub(fs, 'existsSync').callsFake(() => false);
  //   const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/project-x', id: 'b2307a39-e878-458b-bc90-03bc578531d6', asFile: true, path: 'abc', fileName: 'test.docx' } });
  //   Utils.restore(fs.existsSync);
  //   assert.notEqual(actual, true);
  // });

  // it('writeFile called when option --asFile is specified', (done) => {
  //   stubAuth();

  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('/_api/web/GetFileById(') > -1) {
  //       return Promise.resolve('abc');
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso.sharepoint.com';
  //   cmdInstance.action = command.action();

  //   const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => '');
  //   const options: Object = {
  //     debug: false,
  //     id: 'b2307a39-e878-458b-bc90-03bc578531d6',
  //     webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //     asFile: true,
  //     path: '/Users/user/documents',
  //     fileName: 'Test1.docx'
  //   }

  //   cmdInstance.action({ options: options }, () => {
  //     try {
  //       assert(writeFileSyncStub.called)
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore([
  //         request.post,
  //         request.get,
  //         fs.writeFileSync
  //       ]);
  //     }
  //   });
  // });

  // it('writeFile called when option --asFile is specified (debug)', (done) => {
  //   stubAuth();

  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('/_api/web/GetFileById(') > -1) {
  //       return Promise.resolve('abc');
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso.sharepoint.com';
  //   cmdInstance.action = command.action();

  //   const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => '');
  //   const options: Object = {
  //     debug: true,
  //     id: 'b2307a39-e878-458b-bc90-03bc578531d6',
  //     webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //     asFile: true,
  //     path: '/Users/user/documents',
  //     fileName: 'Test1.docx'
  //   }

  //   cmdInstance.action({ options: options }, () => {
  //     try {
  //       assert(writeFileSyncStub.called)
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore([
  //         request.post,
  //         request.get,
  //         fs.writeFileSync
  //       ]);
  //     }
  //   });
  // });

  // it('writeFile not called when option --asFile and path is empty is specified', (done) => {
  //   stubAuth();

  //   sinon.stub(request, 'get').callsFake((opts) => {
  //     if (opts.url.indexOf('/_api/web/GetFileById(') > -1) {
  //       return Promise.resolve('abc');
  //     }

  //     return Promise.reject('Invalid request');
  //   });

  //   auth.site = new Site();
  //   auth.site.connected = true;
  //   auth.site.url = 'https://contoso.sharepoint.com';
  //   cmdInstance.action = command.action();

  //   const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => '');
  //   const options: Object = {
  //     debug: false,
  //     id: 'b2307a39-e878-458b-bc90-03bc578531d6',
  //     webUrl: 'https://contoso.sharepoint.com/sites/project-x',
  //     asFile: true,
  //     fileName: 'Test1.docx'
  //   }

  //   cmdInstance.action({ options: options }, () => {

  //     try {
  //       assert(writeFileSyncStub.notCalled)
  //       done();
  //     }
  //     catch (e) {
  //       done(e);
  //     }
  //     finally {
  //       Utils.restore([
  //         request.post,
  //         request.get,
  //         fs.writeFileSync
  //       ]);
  //     }
  //   });
  // });

  it('fails to update a list item when \'fail me\' values are used', (done) => {

    actualId = 0;
    
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: false, 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      folder: "folder",
      path: "path"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(actualId, 0);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns listItemInstance object when list item is updated with correct values', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    command.allowUnknownOptions();

    let options: any = { 
      debug: true, 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      folder: "folder",
      path: "path",
      contentType: "Document"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(actualId, expectedId);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('attempts to update the listitem with the contenttype of \'Item\' when content type option \'Item\' is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: false, 
      webUrl: 'https://contoso.sharepoint.com/sites/project-y', 
      contentType: 'Item',
      folder: "folder",
      path: "path"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(expectedContentType == actualContentType);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });

  });

  it('attempts to update the listitem with the contenttype of \'Item\' when content type option 0x01 is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: true, 
      webUrl: 'https://contoso.sharepoint.com/sites/project-y', 
      contentType: expectedContentType,
      folder: "folder",
      path: "path"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(expectedContentType == actualContentType);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });

  });


  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.equal(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.FILE_ADD));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com",
        folder: "folder",
        path: "path",
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});