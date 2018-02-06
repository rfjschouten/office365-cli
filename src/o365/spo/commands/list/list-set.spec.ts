import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./list-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.LIST_SET, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  
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
      auth.restoreAuth,
      request.post
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LIST_SET), true);
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
        assert.equal(telemetry.name, commands.LIST_SET);
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
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', title: 'Documents' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified title for list', (done) => {
    const expected = 'List 1';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.Title;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', title: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified baseTemplate for list', (done) => {
    const expected = 100;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.BaseTemplate;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified description for list', (done) => {
    const expected = 'List 1 description';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.Description;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', description: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified templateFeatureId for list', (done) => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.TemplateFeatureId;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified schemaXml for list', (done) => {
    const expected = `<List Title=\'List 1' ID='BE9CE88C-EF3A-4A61-9A8E-F8C038442227'></List>`;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.SchemaXml;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', schemaXml: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowDeletion for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.AllowDeletion;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowDeletion: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowEveryoneViewItems for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.AllowEveryoneViewItems;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowEveryoneViewItems: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified allowMultiResponses for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.AllowMultiResponses;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowMultiResponses: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified contentTypesEnabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ContentTypesEnabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified crawlNonDefaultViews for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.CrawlNonDefaultViews;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', crawlNonDefaultViews: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultContentApprovalWorkflowId for list', (done) => {
    const expected = '00bfea71-de22-43b2-a848-c05709900100';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.DefaultContentApprovalWorkflowId;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultDisplayFormUrl for list', (done) => {
    const expected = '/sites/project-x/List%201/view.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.DefaultDisplayFormUrl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultDisplayFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified defaultEditFormUrl for list', (done) => {
    const expected = '/sites/project-x/List%201/edit.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.DefaultEditFormUrl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultEditFormUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified direction for list', (done) => {
    const expected = 'LTR';
    let actual = '';

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists(guid`) > -1) {
        actual = opts.body.Direction;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified disableGridEditing for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.DisableGridEditing;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableGridEditing: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified draftVersionVisibility for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.DraftVersionVisibility;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified emailAlias for list', (done) => {
    const expected = 'yourname@contoso.onmicrosoft.com';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EmailAlias;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableAssignToEmail for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableAssignToEmail;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAssignToEmail: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableAttachments for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableAttachments;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAttachments: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableDeployWithDependentList for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableDeployWithDependentList;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableDeployWithDependentList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableFolderCreation for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableFolderCreation;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableFolderCreation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableMinorVersions for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableMinorVersions;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableMinorVersions: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableModeration for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableModeration;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableModeration: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enablePeopleSelector for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnablePeopleSelector;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enablePeopleSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableResourceSelector for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableResourceSelector;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableResourceSelector: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableSchemaCaching for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableSchemaCaching;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSchemaCaching: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableSyndication for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableSyndication;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSyndication: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableThrottling for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableThrottling;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableThrottling: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enableVersioning for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnableVersioning;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableVersioning: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified enforceDataValidation for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.EnforceDataValidation;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enforceDataValidation: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified excludeFromOfflineClient for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ExcludeFromOfflineClient;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', excludeFromOfflineClient: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified fetchPropertyBagForListView for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.FetchPropertyBagForListView;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', fetchPropertyBagForListView: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified followable for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.Followable;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', followable: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified forceCheckout for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ForceCheckout;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceCheckout: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified forceDefaultContentType for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ForceDefaultContentType;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceDefaultContentType: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified hidden for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.Hidden;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', hidden: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified includedInMyFilesScope for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.IncludedInMyFilesScope;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', includedInMyFilesScope: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified irmEnabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.IrmEnabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified irmExpire for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.IrmExpire;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmExpire: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified irmReject for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.IrmReject;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmReject: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified isApplicationList for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.IsApplicationList;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', isApplicationList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified listExperienceOptions for list', (done) => {
    const expected = 'NewExperience';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ListExperienceOptions;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified majorVersionLimit for list', (done) => {
    const expected = 34;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.MajorVersionLimit;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: expected, enableVersioning: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified majorWithMinorVersionsLimit for list', (done) => {
    const expected = 20;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.MajorWithMinorVersionsLimit;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorWithMinorVersionsLimit: expected, enableMinorVersions: true, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified multipleDataList for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.MultipleDataList;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', multipleDataList: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified navigateForFormsPages for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.NavigateForFormsPages;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', navigateForFormsPages: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified needUpdateSiteClientTag for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.NeedUpdateSiteClientTag;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', needUpdateSiteClientTag: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified noCrawl for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.NoCrawl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', noCrawl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified onQuickLaunch for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.OnQuickLaunch;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', onQuickLaunch: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified ordered for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.Ordered;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', ordered: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified parserDisabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ParserDisabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', parserDisabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified readOnlyUI for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ReadOnlyUI;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readOnlyUI: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified readSecurity for list', (done) => {
    const expected = 2;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ReadSecurity;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified requestAccessEnabled for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.RequestAccessEnabled;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', requestAccessEnabled: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified restrictUserUpdates for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.RestrictUserUpdates;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', restrictUserUpdates: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified sendToLocationName for list', (done) => {
    const expected = 'SendToLocation';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.SendToLocationName;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', sendToLocationName: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified sendToLocationUrl for list', (done) => {
    const expected = '/sites/project-x/SendToLocation.aspx';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.SendToLocationUrl;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', sendToLocationUrl: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified showUser for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ShowUser;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', showUser: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified useFormsForDisplay for list', (done) => {
    const expected = true;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.UseFormsForDisplay;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', useFormsForDisplay: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified validationFormula for list', (done) => {
    const expected = `IF(fieldName=true);'truetest':'falsetest'`;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ValidationFormula;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', validationFormula: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified validationMessage for list', (done) => {
    const expected = 'Error on field x';
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.ValidationMessage;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', validationMessage: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets specified writeSecurity for list', (done) => {
    const expected = 4;
    let actual = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        actual = opts.body.WriteSecurity;
        return Promise.resolve({ ErrorMessage: null });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: expected, webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert.equal(actual, expected);
        done();
      }
      catch (e) {
        done(e);
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

  it('offers autocomplete for the baseTemplate option', () => {
    const options = (command.options() as CommandOption[]);
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--baseTemplate') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('offers autocomplete for the direction option', () => {
    const options = (command.options() as CommandOption[]);
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--direction') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('configures command types', () => {
    assert.notEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('fails validation if id option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F' } });
    assert(actual);
  });

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: 'true' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F' } });
    assert(actual);
  });

  it('has correct baseTemplate specified', () => {
    const baseTemplateValue = 'DocumentLibrary';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', baseTemplate: baseTemplateValue } });
    assert(actual === true);
  });

  it('fails if non existing baseTemplate specified', () => {
    const baseTemplateValue = 'foo';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', baseTemplate: baseTemplateValue } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the templateFeatureId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the templateFeatureId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', templateFeatureId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the allowDeletion option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowDeletion: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the allowDeletion option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowDeletion: 'true' } });
    assert(actual);
  });

  it('fails validation if the allowEveryoneViewItems option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowEveryoneViewItems: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the allowEveryoneViewItems option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowEveryoneViewItems: 'true' } });
    assert(actual);
  });

  it('fails validation if the allowMultiResponses option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowMultiResponses: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the allowMultiResponses option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', allowMultiResponses: 'true' } });
    assert(actual);
  });

  it('fails validation if the contentTypesEnabled option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the contentTypesEnabled option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', contentTypesEnabled: 'true' } });
    assert(actual);
  });

  it('fails validation if the crawlNonDefaultViews option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', crawlNonDefaultViews: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the crawlNonDefaultViews option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', crawlNonDefaultViews: 'true' } });
    assert(actual);
  });

  it('fails validation if the disableGridEditing option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableGridEditing: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the disableGridEditing option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', disableGridEditing: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableAssignToEmail option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAssignToEmail: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableAssignToEmail option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAssignToEmail: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableAttachments option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAttachments: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableAttachments option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableAttachments: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableDeployWithDependentList option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableDeployWithDependentList: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableDeployWithDependentList option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableDeployWithDependentList: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableFolderCreation option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableFolderCreation: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableFolderCreation option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableFolderCreation: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableMinorVersions option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableMinorVersions: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableMinorVersions option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableMinorVersions: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableModeration option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableModeration: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableModeration option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableModeration: 'true' } });
    assert(actual);
  });

  it('fails validation if the enablePeopleSelector option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enablePeopleSelector: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enablePeopleSelector option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enablePeopleSelector: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableResourceSelector option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableResourceSelector: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableResourceSelector option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableResourceSelector: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableSchemaCaching option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSchemaCaching: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableSchemaCaching option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSchemaCaching: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableSyndication option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSyndication: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableSyndication option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableSyndication: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableThrottling option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableThrottling: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableThrottling option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableThrottling: 'true' } });
    assert(actual);
  });

  it('fails validation if the enableVersioning option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableVersioning: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enableVersioning option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enableVersioning: 'true' } });
    assert(actual);
  });

  it('fails validation if the enforceDataValidation option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enforceDataValidation: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the enforceDataValidation option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', enforceDataValidation: 'true' } });
    assert(actual);
  });

  it('fails validation if the excludeFromOfflineClient option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', excludeFromOfflineClient: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the excludeFromOfflineClient option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', excludeFromOfflineClient: 'true' } });
    assert(actual);
  });

  it('fails validation if the fetchPropertyBagForListView option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', fetchPropertyBagForListView: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the fetchPropertyBagForListView option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', fetchPropertyBagForListView: 'true' } });
    assert(actual);
  });

  it('fails validation if the followable option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', followable: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the followable option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', followable: 'true' } });
    assert(actual);
  });

  it('fails validation if the forceCheckout option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceCheckout: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the forceCheckout option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceCheckout: 'true' } });
    assert(actual);
  });

  it('fails validation if the forceDefaultContentType option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceDefaultContentType: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the forceDefaultContentType option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', forceDefaultContentType: 'true' } });
    assert(actual);
  });

  it('fails validation if the hidden option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', hidden: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the hidden option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', hidden: 'true' } });
    assert(actual);
  });

  it('fails validation if the includedInMyFilesScope option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', includedInMyFilesScope: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the includedInMyFilesScope option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', includedInMyFilesScope: 'true' } });
    assert(actual);
  });

  it('fails validation if the irmEnabled option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmEnabled: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the irmEnabled option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmEnabled: 'true' } });
    assert(actual);
  });

  it('fails validation if the irmExpire option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmExpire: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the irmExpire option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmExpire: 'true' } });
    assert(actual);
  });

  it('fails validation if the irmReject option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmReject: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the irmReject option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', irmReject: 'true' } });
    assert(actual);
  });

  it('fails validation if the isApplicationList option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', isApplicationList: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the isApplicationList option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', isApplicationList: 'true' } });
    assert(actual);
  });

  it('fails validation if the multipleDataList option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', multipleDataList: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the multipleDataList option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', multipleDataList: 'true' } });
    assert(actual);
  });

  it('fails validation if the navigateForFormsPages option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', navigateForFormsPages: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the navigateForFormsPages option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', navigateForFormsPages: 'true' } });
    assert(actual);
  });

  it('fails validation if the needUpdateSiteClientTag option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', needUpdateSiteClientTag: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the needUpdateSiteClientTag option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', needUpdateSiteClientTag: 'true' } });
    assert(actual);
  });

  it('fails validation if the noCrawl option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', noCrawl: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the noCrawl option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', noCrawl: 'true' } });
    assert(actual);
  });

  it('fails validation if the onQuickLaunch option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', onQuickLaunch: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the onQuickLaunch option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', onQuickLaunch: 'true' } });
    assert(actual);
  });

  it('fails validation if the ordered option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', ordered: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the ordered option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', ordered: 'true' } });
    assert(actual);
  });

  it('fails validation if the parserDisabled option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', parserDisabled: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the parserDisabled option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', parserDisabled: 'true' } });
    assert(actual);
  });

  it('fails validation if the readOnlyUI option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readOnlyUI: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the readOnlyUI option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readOnlyUI: 'true' } });
    assert(actual);
  });

  it('fails validation if the requestAccessEnabled option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', requestAccessEnabled: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the requestAccessEnabled option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', requestAccessEnabled: 'true' } });
    assert(actual);
  });

  it('fails validation if the restrictUserUpdates option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', restrictUserUpdates: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the restrictUserUpdates option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', restrictUserUpdates: 'true' } });
    assert(actual);
  });

  it('fails validation if the showUser option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', showUser: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the showUser option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', showUser: 'true' } });
    assert(actual);
  });

  it('fails validation if the useFormsForDisplay option is not a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', useFormsForDisplay: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the useFormsForDisplay option is a valid Boolean', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', useFormsForDisplay: 'true' } });
    assert(actual);
  });

  it('fails validation if the defaultContentApprovalWorkflowId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the defaultContentApprovalWorkflowId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', defaultContentApprovalWorkflowId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails if non existing draftVersionVisibility specified', () => {
    const draftVersionValue = 'NonExistingDraftVersionVisibility';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: draftVersionValue } });
    assert.notEqual(actual, true);
  });

  it('has correct draftVersionVisibility specified', () => {
    const draftVersionValue = 'Approver';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', draftVersionVisibility: draftVersionValue } });
    assert(actual === true);
  });

  it('fails if emailAlias specified, but enableAssignToEmail is not true', () => {
    const emailAliasValue = 'yourname@contoso.onmicrosoft.com';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: emailAliasValue } });
    assert.equal(actual, `emailAlias could not be set if enableAssignToEmail is not set to true. Please set enableAssignToEmail.`);
  });

  it('has correct emailAlias and enableAssignToEmail values specified', () => {
    const emailAliasValue = 'yourname@contoso.onmicrosoft.com';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', emailAlias: emailAliasValue, enableAssignToEmail: 'true' } });
    assert(actual === true);
  });

  it('fails if non existing direction specified', () => {
    const directionValue = 'abc';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: directionValue } });
    assert.notEqual(actual, true);
  });

  it('has correct direction specified', () => {
    const directionValue = 'LTR';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', direction: directionValue } });
    assert(actual === true);
  });

  it('fails if majorVersionLimit specified, but enableVersioning is not true', () => {
    const majorVersionLimitValue = 20;
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: majorVersionLimitValue } });
    assert.equal(actual, `majorVersionLimit option is only valid in combination with enableVersioning.`);
  });

  it('has correct majorVersionLimit and enableVersioning values specified', () => {
    const majorVersionLimitValue = 20;
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: majorVersionLimitValue, enableVersioning: 'true' } });
    assert(actual === true);
  });
  
  it('fails if majorWithMinorVersionsLimit specified, but enableModeration is not true', () => {
    const majorWithMinorVersionLimitValue = 20;
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorWithMinorVersionsLimit: majorWithMinorVersionLimitValue } });
    assert.equal(actual, `majorWithMinorVersionsLimit option is only valid in combination with enableMinorVersions or enableModeration.`);
  });

  it('has correct majorVersionLimit and enableVersioning values specified', () => {
    const majorVersionLimitValue = 20;
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', majorVersionLimit: majorVersionLimitValue, enableVersioning: 'true' } });
    assert(actual === true);
  });

  it('fails if non existing readSecurity specified', () => {
    const readSecurityValue = 5;
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: readSecurityValue } });
    assert.notEqual(actual, true);
  });

  it('has correct readSecurity specified', () => {
    const readSecurityValue = 2;
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', readSecurity: readSecurityValue } });
    assert(actual === true);
  });

  it('fails if non existing listExperienceOptions specified', () => {
    const listExperienceValue = 'NonExistingExperience';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: listExperienceValue } });
    assert.notEqual(actual, true);
  });

  it('has correct listExperienceOptions specified', () => {
    const listExperienceValue = 'NewExperience';
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', listExperienceOptions: listExperienceValue } });
    assert(actual === true);
  });

  it('fails if non existing readSecurity specified', () => {
    const writeSecurityValue = 5;
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: writeSecurityValue } });
    assert.notEqual(actual, true);
  });

  it('has correct direction specified', () => {
    const writeSecurityValue = 4;
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '3EA5A977-315E-4E25-8B0F-E4F949BF6B8F', writeSecurity: writeSecurityValue } });
    assert(actual === true);
  });

  it('returns listInstance object when list is added with correct values', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/lists`) > -1) {
        return Promise.resolve(
          {
            "AllowContentTypes": true,
            "BaseTemplate": 100,
            "BaseType": 1,
            "ContentTypesEnabled": false,
            "CrawlNonDefaultViews": false,
            "Created": null,
            "CurrentChangeToken": null,
            "CustomActionElements": null,
            "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000",
            "DefaultItemOpenUseListSetting": false,
            "Description": "",
            "Direction": "none",
            "DocumentTemplateUrl": null,
            "DraftVersionVisibility": 0,
            "EnableAttachments": false,
            "EnableFolderCreation": true,
            "EnableMinorVersions": false,
            "EnableModeration": false,
            "EnableVersioning": false,
            "EntityTypeName": "Documents",
            "ExemptFromBlockDownloadOfNonViewableFiles": false,
            "FileSavePostProcessingEnabled": false,
            "ForceCheckout": false,
            "HasExternalDataSource": false,
            "Hidden": false,
            "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
            "ImagePath": null,
            "ImageUrl": null,
            "IrmEnabled": false,
            "IrmExpire": false,
            "IrmReject": false,
            "IsApplicationList": false,
            "IsCatalog": false,
            "IsPrivate": false,
            "ItemCount": 69,
            "LastItemDeletedDate": null,
            "LastItemModifiedDate": null,
            "LastItemUserModifiedDate": null,
            "ListExperienceOptions": 0,
            "ListItemEntityTypeFullName": null,
            "MajorVersionLimit": 0,
            "MajorWithMinorVersionsLimit": 0,
            "MultipleDataList": false,
            "NoCrawl": false,
            "ParentWebPath": null,
            "ParentWebUrl": null,
            "ParserDisabled": false,
            "ServerTemplateCanCreateFolders": true,
            "TemplateFeatureId": null,
            "Title": "List 1"
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, title: 'List 1', baseTemplate: 'GenericList', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ 
          AllowContentTypes: true,
          BaseTemplate: 100,
          BaseType: 1,
          ContentTypesEnabled: false,
          CrawlNonDefaultViews: false,
          Created: null,
          CurrentChangeToken: null,
          CustomActionElements: null,
          DefaultContentApprovalWorkflowId: '00000000-0000-0000-0000-000000000000',
          DefaultItemOpenUseListSetting: false,
          Description: '',
          Direction: 'none',
          DocumentTemplateUrl: null,
          DraftVersionVisibility: 0,
          EnableAttachments: false,
          EnableFolderCreation: true,
          EnableMinorVersions: false,
          EnableModeration: false,
          EnableVersioning: false,
          EntityTypeName: 'Documents',
          ExemptFromBlockDownloadOfNonViewableFiles: false,
          FileSavePostProcessingEnabled: false,
          ForceCheckout: false,
          HasExternalDataSource: false,
          Hidden: false,
          Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
          ImagePath: null,
          ImageUrl: null,
          IrmEnabled: false,
          IrmExpire: false,
          IrmReject: false,
          IsApplicationList: false,
          IsCatalog: false,
          IsPrivate: false,
          ItemCount: 69,
          LastItemDeletedDate: null,
          LastItemModifiedDate: null,
          LastItemUserModifiedDate: null,
          ListExperienceOptions: 0,
          ListItemEntityTypeFullName: null,
          MajorVersionLimit: 0,
          MajorWithMinorVersionsLimit: 0,
          MultipleDataList: false,
          NoCrawl: false,
          ParentWebPath: null,
          ParentWebUrl: null,
          ParserDisabled: false,
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: null,
          Title: 'List 1'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
    
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
    assert(find.calledWith(commands.LIST_SET));
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
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
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