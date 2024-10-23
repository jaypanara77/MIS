import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IMisEventVersionProps } from './IMisEventVersionProps';

interface IVersionInfo {
  version: string;
  attachmentLink?: string;
}

interface AttachmentLink {
  fileName: string;
  fileUrl: string;
  versionLabel: string;
  versionNumber: string; // Use the custom 'Version_number'
}

export interface IMisEventVersionState {
  versionHistory: IVersionInfo[];
  attachmentLinks: AttachmentLink[];
  loading: boolean;
  error: string | null;
}

export default class MisEventVersion extends React.Component<IMisEventVersionProps, IMisEventVersionState> {
  constructor(props: IMisEventVersionProps) {
    super(props);

    this.state = {
      versionHistory: [],
      attachmentLinks: [],
      loading: true,
      error: null,
    };
  }

  public componentDidMount(): void {
    this.getItemVersionHistory();
  }

  // Fetch version history and attachments
  private async getItemVersionHistory(): Promise<void> {
    const ndcCode = new URLSearchParams(window.location.search).get('NDCCode');
    if (!ndcCode) {
      this.setState({ loading: false, error: 'NDCCode not found in the URL' });
      return;
    }

    try {
      // Fetch list item ID using NDCCode
      const listItemId = await this.getListItemId(ndcCode);
      if (listItemId === 0) {
        this.setState({ loading: false, error: 'List item not found' });
        return;
      }

      // Get version history from the list
      const versionHistoryUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('MIS_Upload_File')/items(${listItemId})/versions`;
      const response = await this.props.spHttpClient.get(versionHistoryUrl, SPHttpClient.configurations.v1);
      const versionHistoryData = await response.json();

      // Fetch the attachments from document set (MIS_Attachment)
      await this.getDocumentSetAttachments(ndcCode);

      // Map version history and compare with the attachments' "Version_number"
      const versionHistory: IVersionInfo[] = versionHistoryData.value.map((version: any) => {
        const matchingAttachment = this.state.attachmentLinks.find((att: AttachmentLink) => att.versionNumber === version.VersionLabel); // Compare version number
        return {
          version: version.VersionLabel,
          attachmentLink: matchingAttachment ? matchingAttachment.fileUrl : 'No attachments'
        };
      });

      // Update the state with version history and attachments
      this.setState({ versionHistory, loading: false });
    } catch (error) {
      console.error('Error retrieving version history or attachments:', error);
      this.setState({ loading: false, error: 'Error retrieving data.' });
    }
  }

  // Fetch list item ID by NDCCode
  private async getListItemId(ndcCode: string): Promise<number> {
    const listUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('MIS_Upload_File')/items?$filter=NDCCode eq '${ndcCode}'`;
    const response = await this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
    const data = await response.json();
    return data.value[0]?.Id ?? -1;
  }

  // Fetch document set attachments and their corresponding version numbers
  private async getDocumentSetAttachments(ndcCode: string): Promise<void> {
    if (!ndcCode) {
      this.setState({ loading: false, error: 'NDCCode not found in the URL' });
      return;
    }
  
    try {
      // Construct the API URL to get files from the document set (MIS_Attachment)
      const docSetUrl = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('MIS_Attachment/${ndcCode}')/Files?$expand=ListItemAllFields&$select=Name,ServerRelativeUrl,ListItemAllFields/Version_number,ListItemAllFields/ID,UIVersionLabel`;
  
      // Make the HTTP request to get the files in the document set
      const response: SPHttpClientResponse = await this.props.spHttpClient.get(docSetUrl, SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
  
      const files = await response.json();
  
      // Log the response to check its structure
      console.log('Files response:', files);
  
      // Check if the response has the expected structure
      if (files.value && Array.isArray(files.value)) {
        // Extract file names, URLs, UIVersionLabel, and custom Version_number from ListItemAllFields
        const attachmentLinks = files.value.map((file: any) => {
          // Log each file to understand its structure
          console.log('File:', file);
          return {
            fileName: file.Name,
            fileUrl: `${this.props.siteUrl}${file.ServerRelativeUrl}`, // Construct the full URL
            versionLabel: file.UIVersionLabel || 'N/A', // Handle undefined values for UIVersionLabel
            versionNumber: file.ListItemAllFields?.Version_number ?? 'N/A', // Fetch custom 'Version_number' column value
          };
        });
  
        // Update state with the attachment links
        this.setState({ attachmentLinks, loading: false });
      } else {
        this.setState({ loading: false, error: 'No files found in the document set.' });
      }
    } catch (error) {
      console.error('Error fetching document set attachments:', error);
      this.setState({ loading: false, error: 'Error fetching attachments. Please try again later.' });
    }
  }
  

  public render(): React.ReactElement<IMisEventVersionProps> {
    const { loading, error, versionHistory } = this.state;

    return (
      <div>
        <h2>Version History for NDCCode</h2>
        {loading && <p>Loading version history...</p>}
        {error && <p>{error}</p>}
        <ul>
          {versionHistory.map((versionInfo, index) => (
            <li key={index}>
              Version {versionInfo.version}
              {versionInfo.attachmentLink !== 'No attachments' ? (
                <span>
                  {' - '}<a href={versionInfo.attachmentLink} target="_blank" rel="noopener noreferrer">
                    View Attachment
                  </a>
                </span>
              ) : (
                ' - No attachments'
              )}
            </li>
          ))}
        </ul>
      </div>
    );
  }
}
