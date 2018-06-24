import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import * as fs from 'fs';
import * as path from 'path';
import { ContextInfo } from '../../spo';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folder: string;
  path: string;
  contentType?: string;
  checkOut?: boolean;
  checkInComment?: string;
  approve?: boolean;
  approveComment?: string;
  publish?: boolean;
  publishComment?: string;
}

// interface FieldValue {
//   ErrorMessage: string;
//   FieldName: string;
//   FieldValue: any;
//   HasException: boolean;
//   ItemId: number;
// }

class SpoFileAddCommand extends SpoCommand {
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get name(): string {
    return commands.FILE_ADD;
  }

  public get description(): string {
    return 'Upload file to the specified folder';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.contentType = (!(!args.options.contentType)).toString();
    telemetryProps.checkOut = args.options.checkOut || false;
    telemetryProps.checkInComment = (!(!args.options.checkInComment)).toString();
    telemetryProps.approve = args.options.approve || false;
    telemetryProps.approveComment = (!(!args.options.approveComment)).toString();
    telemetryProps.publish = args.options.publish || false;
    telemetryProps.publishComment = (!(!args.options.publishComment)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let requestDigest: string = '';
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = path.basename(fullPath);
    let fileExists: boolean = true;
    let contentTypeName: string = '';

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving request digest...`);
        }

        return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
      })
      .then((res: ContextInfo): request.RequestPromise | Promise<string> => {
        requestDigest = res.FormDigestValue;
        
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          })
        };
        
        return request.get(requestOptions);  
      })
      .catch((err: string): request.RequestPromise => {
          if (this.debug) {
            cmd.log('Folder does not exist so create it')
            cmd.log('');
          }
          const folderBody: string = `{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '${encodeURIComponent(args.options.folder)}'}`;
          const bodyLength: number = folderBody.length;
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/folders`,
            body: folderBody,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-RequestDigest': requestDigest,
              accept: 'application/json;odata=verbose',
              'content-type': 'application/json;odata=verbose',
              'content-length': bodyLength
            })
          };
  
          return request.post(requestOptions);
      })
      .then((res: ContextInfo): request.RequestPromise | Promise<void> => {
        if (!args.options.checkOut) {
          // no checkOut needed so go on uploading file
          return Promise.resolve();
        }
        // if checkOut then check if file already exists, otherwise it can't be checked out
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          })
        };
        
        return request.get(requestOptions);
      })
      .catch((err: string): Promise<void> => {
        fileExists = false;
        return Promise.resolve();
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (fileExists && args.options.checkOut) {
          if (this.verbose) {
            cmd.log(`Checkout file ${fileName}`);
          }
          // checkout the existing file
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')/CheckOut()`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-RequestDigest': requestDigest,
            })
          };
          
          return request.post(requestOptions);
        }

        return Promise.resolve();
      })
      .then((res: any): request.RequestPromise => {
        if (this.verbose) {
          cmd.log(`Upload file to site ${args.options.webUrl}...`);
        }

        const fileBody: Buffer = fs.readFileSync(fullPath);
        const bodyLength: number = fileBody.byteLength;

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files/Add(url='${encodeURIComponent(fileName)}', overwrite=true)`,
          body: fileBody,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'X-RequestDigest': requestDigest,
            'content-length':bodyLength
          })
        };

        return request.post(requestOptions);
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (args.options.contentType) {
          if (this.verbose) {
            cmd.log(`Getting list id in order to get its available content types afterwards...`);
          }

          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')/ListItemAllFields/ParentList?$Select=Id`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'accept': 'application/json;odata=nometadata'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.get(requestOptions);
        }

        return Promise.resolve();
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (args.options.contentType) {
          if (this.debug) {
            cmd.log('list id response...');
            cmd.log(res.Id);
          }
          
          if (this.verbose) {
            cmd.log(`Getting content types for list...`);
          }

          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/lists('${res.Id}')/contenttypes?$select=Name,Id`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'accept': 'application/json;odata=nometadata'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.get(requestOptions);
        }

        return Promise.resolve();
      })
      .then((response: any): request.RequestPromise | Promise<void> => {
        if (args.options.contentType) {

          if (this.debug) {
            cmd.log('content type lookup response...');
            cmd.log(response);
          }

          const foundContentType = response.value.filter((ct: any) => {
            const contentTypeMatch: boolean = ct.Id.StringValue === args.options.contentType || ct.Name === args.options.contentType;

            if (this.debug) {
              cmd.log(`Checking content type value [${ct.Name}]: ${contentTypeMatch}`);
            }

            return contentTypeMatch;
          });

          if (this.debug) {
            cmd.log('content type filter output...');
            cmd.log(foundContentType);
          }

          if (foundContentType.length > 0) {
            contentTypeName = foundContentType[0].Name;
          }

          // After checking for content types, throw an error if the name is blank
          if (!contentTypeName || contentTypeName === '') {
            return Promise.reject(`Specified content type '${args.options.contentType}' doesn't exist on the target list`);
          }

          if (this.debug) {
            cmd.log(`using content type name: ${contentTypeName}`);
          }
        }

        if (this.verbose) {
          cmd.log(`Add List Item values for file ${fileName}`);
        }

        const requestBody: any = {
          formValues: this.mapRequestBody(args.options),
          bNewDocumentUpdate: true
        };

        if (args.options.contentType && contentTypeName !== '') {
          if (this.debug) {
            cmd.log(`Specifying content type name [${contentTypeName}] in request body`);
          }

          requestBody.formValues.push({
            FieldName: 'ContentType',
            FieldValue: contentTypeName
          });
        }

        cmd.log(JSON.stringify(requestBody));

        // Checkin the existing file with given comment
        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')/ListItemAllFields/ValidateUpdateListItem()`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: requestBody,
          json: true
        };

        return request.post(requestOptions);
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (fileExists && args.options.checkOut) {
          if (this.verbose) {
            cmd.log(`Checkin file ${fileName}`);
          }
          // Checkin the existing file with given comment
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')/CheckIn(comment='${encodeURIComponent(args.options.checkInComment || '')}',checkintype=0)`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-RequestDigest': requestDigest,
            })
          };
          
          return request.post(requestOptions);
        }

        return Promise.resolve();
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (args.options.approve) {
          if (this.verbose) {
            cmd.log(`Approve file ${fileName}`);
          }
          // Checkin the existing file with given comment
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')/approve(comment='${encodeURIComponent(args.options.approveComment || '')}')`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-RequestDigest': requestDigest,
            })
          };
          
          return request.post(requestOptions);
        }

        return Promise.resolve();
      })
      .then((res: any): request.RequestPromise | Promise<void> => {
        if (args.options.publish) {
          if (this.verbose) {
            cmd.log(`Publish file ${fileName}`);
          }
          // Checkin the existing file with given comment
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')/publish(comment='${encodeURIComponent(args.options.publishComment || '')}')`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-RequestDigest': requestDigest,
            })
          };
          
          return request.post(requestOptions);
        }

        return Promise.resolve();
      })
      .then((file: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(file);
          cmd.log('');
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = [];
    //const requestBody1: any = {};
    const excludeOptions: string[] = [
      'webUrl',
      'folder',
      'path',
      'contentType',
      'checkOut',
      'checkInComment',
      'approve',
      'approveComment',
      'publish',
      'publishComment',
      'debug',
      'verbose'
    ];

    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
          requestBody.push({FieldName: key, FieldValue: (<any>options)[key].toString() });
      }
    });

    return requestBody;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-w, --webUrl <webUrl>',
        description: 'The URL of the site where the file should be uploaded to'
      },
      {
        option: '-f, --folder <folder>',
        description: 'Server-relative URL to the folder where the file should be uploaded'
      },
      {
        option: '-p, --path <path>',
        description: 'local path to the file to upload'
      },
      {
        option: '-c, --contentType [contentType]',
        description: 'Content type name or ID to assign to the file'
      },
      {
        option: '--checkOut',
        description: 'If versioning is enabled, this will check out the file first if it exists, upload the file, then check it in again'
      },
      {
        option: '--checkInComment [checkInComment]',
        description: 'Comment to set when checking the file in'
      },
      {
        option: '--approve',
        description: 'Will automatically approve the uploaded file'
      },
      {
        option: '--approveComment [approveComment]',
        description: 'Comment to set when approving the file'
      },
      {
        option: '--publish',
        description: 'Will automatically publish the uploaded file'
      },
      {
        option: '--publishComment [publishComment]',
        description: 'Comment to set when publishing the file'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (args.options.path && !fs.existsSync(args.options.path)) {
        return 'Specified path of the file to add does not exits';
      }

      // if (args.options.id) {
      //   if (!Utils.isValidGuid(args.options.id)) {
      //     return `${args.options.id} is not a valid GUID`;
      //   }
      // }

      // if (args.options.id && args.options.url) {
      //   return 'Specify id or url, but not both';
      // }

      // if (!args.options.id && !args.options.url) {
      //   return 'Specify id or url, one is required';
      // }

      // if (args.options.asFile && !args.options.path) {
      //   return 'The path should be specified when the --asFile option is used';
      // }

      // if (args.options.path && !fs.existsSync(path.dirname(args.options.path))) {
      //   return 'Specified path where to save the file does not exits';
      // }

      // if (args.options.asFile) {
      //   if (args.options.asListItem || args.options.asString) {
      //     return 'Specify to retrieve the file either as file, list item or string but not multiple';
      //   }
      // }

      // if (args.options.asListItem) {
      //   if (args.options.asFile || args.options.asString) {
      //     return 'Specify to retrieve the file either as file, list item or string but not multiple';
      //   }
      // }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.
  
  Remarks:
  
    To get a file, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Get file properties for file with id ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6'

    Get contents of the file with id ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asString

    Get list item properties for file with id
    ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asListItem   

    Save file with id ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')} located
    in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')} to local file
    ${chalk.grey('/Users/user/documents/SavedAsTest1.docx')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --id 'b2307a39-e878-458b-bc90-03bc578531d6' --asFile --path /Users/user/documents/SavedAsTest1.docx
    
    Return file properties for file with server-relative url
    ${chalk.grey('/sites/project-x/documents/Test1.docx')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx'

    Return file as string for file with server-relative url
    ${chalk.grey('/sites/project-x/documents/Test1.docx')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asString

    Return list item properties for file with server-relative url
    ${chalk.grey('/sites/project-x/documents/Test1.docx')} located in site
    ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asListItem   

    Save file with server-relative url ${chalk.grey('/sites/project-x/documents/Test1.docx')}
    located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
    to local file ${chalk.grey('/Users/user/documentsSavedAsTest1.docx')}
      ${chalk.grey(config.delimiter)} ${commands.FILE_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --url '/sites/project-x/documents/Test1.docx' --asFile --path /Users/user/documents/SavedAsTest1.docx
      `);
  }
}

module.exports = new SpoFileAddCommand();