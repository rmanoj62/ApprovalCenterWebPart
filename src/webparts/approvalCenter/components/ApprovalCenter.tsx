import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { IApprovalCenterProps } from './IApprovalCenterProps';

interface IApprovalItem {
  Id: number;
  ApplicationName: string;
  Status: string;
  ApprovalEmail: string;
  ApprovalLink: string;
}

const ApprovalCenter: React.FC<IApprovalCenterProps> = (props) => {

  const [items, setItems] = React.useState<IApprovalItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);

  const loadData = async (): Promise<void> => {

    try {

      const url =
        props.siteUrl +
        "/_api/web/lists/getbytitle('Approvalcenter')/items?$select=Id,ApplicationName,Status,ApprovalEmail,ApprovalLink";

      const response = await props.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      const result = await response.json();

      const currentUser =
        props.currentUserEmail.toLowerCase();

      const filteredItems = result.value.filter((item: IApprovalItem) => {

        if (!item.ApprovalEmail) {
          return false;
        }

        const emails = item.ApprovalEmail
          .split(';')
          .map(function (email: string) {
            return email.trim().toLowerCase();
          });

        return emails.indexOf(currentUser) > -1;
      });

      setItems(filteredItems);
    }
    catch (error) {
      console.log(error);
    }
    finally {
      setLoading(false);
    }
  };

  React.useEffect(() => {
    void loadData();
  }, []);

  return (
    <div
      style={{
        backgroundColor: '#ffffff',
        borderRadius: '8px',
        padding: '20px',
        boxShadow: '0 2px 8px rgba(0,0,0,0.08)'
      }}
    >

      {/* Title */}

      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          marginBottom: '20px'
        }}
      >
        <Icon
          iconName="CompletedSolid"
          style={{
            color: '#107C10',
            fontSize: '22px',
            marginRight: '10px'
          }}
        />

        <span
          style={{
            fontSize: '22px',
            fontWeight: 600,
            color: '#323130'
          }}
        >
          Approval Center
        </span>
      </div>

      {loading && <div>Loading...</div>}

      {!loading && items.length === 0 && (
        <div>No records found.</div>
      )}

      {!loading && items.length > 0 && (

        <table
          style={{
            width: '100%',
            borderCollapse: 'collapse'
          }}
        >
          <thead>
            <tr>

              <th style={headerStyle}>
                <Icon
                  iconName="AppIconDefault"
                  style={headerIconStyle}
                />
                Application Name
              </th>

              <th style={headerStyle}>
                <Icon
                  iconName="StatusCircleCheckmark"
                  style={headerIconStyle}
                />
                Status
              </th>

              <th style={headerStyle}>
                <Icon
                  iconName="Link"
                  style={headerIconStyle}
                />
                Approval Link
              </th>

            </tr>
          </thead>

          <tbody>

            {items.map(function (item) {

              return (
                <tr key={item.Id}>
                  <td style={cellStyle}>
                    {item.ApplicationName}
                  </td>

                  <td style={cellStyle}>
                    {item.Status}
                  </td>

                  <td style={cellStyle}>

                    <a
                      href={item.ApprovalLink}
                      target="_blank"
                      rel="noopener noreferrer"
                      style={{
                        backgroundColor: '#107C10',
                        color: '#ffffff',
                        padding: '6px 14px',
                        borderRadius: '4px',
                        textDecoration: 'none',
                        fontSize: '13px'
                      }}
                    >
                      Open
                    </a>

                  </td>
                </tr>
              );
            })}

          </tbody>

        </table>

      )}

    </div>
  );
};

const headerStyle = {
  border: '1px solid #e1dfdd',
  padding: '12px',
  backgroundColor: '#f3f2f1',
  color: '#323130',
  textAlign: 'left' as const,
  fontWeight: 600
};

const headerIconStyle = {
  color: '#107C10',
  fontSize: '12px',
  marginRight: '6px'
};

const cellStyle = {
  border: '1px solid #e1dfdd',
  padding: '12px'
};

export default ApprovalCenter;