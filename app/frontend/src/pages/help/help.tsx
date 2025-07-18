import React from "react";
import styles from "./help.module.css";

const Help: React.FC = () => {
  return (
    <div className={styles.helpContainer}>
      <h1 className={styles.helpTitle}>Help & Support</h1>
      <p className={styles.helpSubtitle}>
        Access documentation and support resources
      </p>
      
      <div className={styles.linkCard}>
        <div className={styles.linkIcon}>📚</div>
        <h3>Documentation</h3>
        <p>Complete guide to using the internship portal</p>
        <a
          href="https://s2as.sharepoint.com/:w:/r/sites/InternshipPortal/_layouts/15/Doc2.aspx?action=edit&sourcedoc=%7Bece22c69-3794-4a9c-9b25-6a0e6acfbfb6%7D&wdOrigin=TEAMS-MAGLEV.teamsSdk_ns.rwc&wdExp=TEAMS-TREATMENT&wdhostclicktime=1752849287618&web=1"
          target="_blank"
          rel="noopener noreferrer"
          className={styles.primaryLink}
        >
          Open Documentation
        </a>
      </div>
    </div>
  );
};

export default Help;