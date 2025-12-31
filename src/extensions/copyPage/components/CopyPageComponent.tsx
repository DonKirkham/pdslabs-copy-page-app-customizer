import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import CopyPageDialog from './CopyPageDialog';
import { useBoolean } from '@fluentui/react-hooks';
import { ISPFXContext } from '@pnp/sp';
import styles from './CopyPage.module.scss';

export interface ICopyPageComponentProps {
  context: ISPFXContext;
  pageName: string;
  pageUrl: string;
  siteUrl: string;
}

const CopyPageComponent: React.FC<ICopyPageComponentProps> = (props) => {
  const [hideDialog, { toggle }] = useBoolean(true);
  const [currentUrl, setCurrentUrl] = React.useState(window.location.href);
  
  // Monitor URL changes to detect mode switches
  React.useEffect(() => {
    const checkUrl = setInterval(() => {
      if (window.location.href !== currentUrl) {
        setCurrentUrl(window.location.href);
      }
    }, 500); // Check every 500ms

    return () => clearInterval(checkUrl);
  }, [currentUrl]);
  
  // Check if page is in edit mode - runs on every render
  const isEditMode = currentUrl.indexOf('Mode=Edit') > -1 || 
                     currentUrl.indexOf('/_layouts/15/') > -1;

  console.log('CopyPageComponent - URL:', currentUrl);
  console.log('CopyPageComponent - isEditMode:', isEditMode);

  // Only show button when in edit mode
  if (!isEditMode) {
    return null;
  }

  return (
    <div className={styles.appContainer}>

      {/* Button to toggle the Copy Page Dialog */}
      <DefaultButton
        className={styles.appButton}
        onClick={toggle}
        iconProps={{ iconName: 'Copy' }} // Add the "Copy" icon
      >
        Copy Page
      </DefaultButton>

      {/* Copy Page Dialog */}
      <CopyPageDialog hidden={hideDialog} onDismiss={toggle} {...props} />

    </div>
  );
};

export default CopyPageComponent;