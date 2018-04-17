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
import Auth from '../../../../Auth';
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
  values?: string;
}

class SpoFileAddCommand extends SpoCommand {
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
    telemetryProps.values = (!(!args.options.values)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken: string = '';
    let requestDigest: string = '';
    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = path.basename(fullPath);
    let fileExists: boolean = true;

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
        if (args.options.values) {
          if (this.verbose) {
            cmd.log(`Add List Item values for file ${fileName}`);
          }

          let body = this.buildBody(args.options.values);
          cmd.log(JSON.stringify(body));

          // Checkin the existing file with given comment
          const requestOptions: any = {
            url: `${args.options.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(args.options.folder)}')/Files('${encodeURIComponent(fileName)}')/ListItemAllFields`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-RequestDigest': requestDigest,
              'X-HTTP-Method': 'PATCH',
              'IF-MATCH': '*',
              'content-type': 'application/json;odata=verbose',
              accept: 'application/json;odata=verbose'
            }),
            body: JSON.stringify(body)
          };
          
          return request.post(requestOptions);
        }

        return Promise.resolve();
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

  /**
   * splits a values string separated by ; and for each value it will get its key/values seperated by =
   * example values string: InternalFieldName1=value1;InternalFieldName2=value2;InternalFieldName3=value3
   * above string should output a body object {"InternalFieldName1":"value1","InternalFieldName2":"value2","InternalFieldName3":"value3"}
   */
  private buildBody(values: string): any {
    let body = new Map<string, any>();
    
    const valuesArray: string[] = values.split(';');
    valuesArray.map(property => {
      property.trim();
      const propertyArray: string[] = property.split('=');
      body.set(propertyArray[0], propertyArray[1]);
    });
    //const itemArray: string[] = valuesArray.split(';').map(v => v.trim());
    
    // if (values.length > 0) {
    //   if (values.indexOf(';') > 0) {
    //     let valuesArray = values.split(';');
    //     valuesArray.forEach(value => {
    //       if (value.indexOf('=') > 0) {
    //         let valueArray = value.split('=');
    //         //valueArray.forEach(valueElements => {
    //           body.set(valueArray[0], valueArray[1]);
    //         //});
    //       }
    //     });
    //   }
    //   else {
    //     if (values.indexOf('=') > 0){
    //       let valueElements = values.split('=');
    //       body.set(valueElements[0], valueElements[1]);
    //     }
    //   }
    // }

    var itemPayload: any = {};
    itemPayload['__metadata'] = {'type':'SP.ListItem'};
    
    body.forEach((value: any, key: string) => {
      itemPayload[key] = value;
    });

    return itemPayload;
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
      },
      {
        option: '--values [values]',
        description: 'this command should allow using unknown properties. Each property corresponds to the list item field that should be set when uploading the file.'
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