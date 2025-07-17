import React, { useState, useEffect, useRef, RefObject } from "react";
import { Outlet, NavLink, Link } from "react-router-dom";
import { useTranslation } from "react-i18next";
import styles from "./Layout.module.css";

import { useMsal } from "@azure/msal-react"; // Add this import if using MSAL
import { useLogin } from "../../authConfig";
import { LoginButton } from "../../components/LoginButton";
import { IconButton } from "@fluentui/react";

const Layout = () => {
    const { t } = useTranslation();
    const [menuOpen, setMenuOpen] = useState(false);
    const menuRef: RefObject<HTMLDivElement> = useRef(null);

    // User info (MSAL)
    const { accounts } = useMsal ? useMsal() : { accounts: [] };
    const user = accounts && accounts.length > 0 ? accounts[0] : null;

    // Date/time state
    const [dateTime, setDateTime] = useState(new Date());
    useEffect(() => {
        const interval = setInterval(() => setDateTime(new Date()), 1000);
        return () => clearInterval(interval);
    }, []);

    const toggleMenu = () => {
        setMenuOpen(!menuOpen);
    };

    const handleClickOutside = (event: MouseEvent) => {
        if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
            setMenuOpen(false);
        }
    };

    useEffect(() => {
        if (menuOpen) {
            document.addEventListener("mousedown", handleClickOutside);
        } else {
            document.removeEventListener("mousedown", handleClickOutside);
        }
        return () => {
            document.removeEventListener("mousedown", handleClickOutside);
        };
    }, [menuOpen]);

    // Format date and time with better styling
    const formatDate = (date: Date) => {
        return date.toLocaleDateString('en-US', {
            weekday: 'short',
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
    };

    const formatTime = (date: Date) => {
        return date.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: false
        });
    };

    // Get user initials for avatar
    const getUserInitials = (user: any) => {
        if (!user) return "G";
        const name = user.name || user.username || "Guest";
        return name.split(' ').map((n: string) => n[0]).join('').toUpperCase().slice(0, 2);
    };

    const [tokenUsage, setTokenUsage] = useState(0); // 0-100 percent

    return (
        <div className={styles.layout}>
            <header className={styles.header} role={"banner"}>
                <div className={styles.headerContainer} ref={menuRef}>
                    {/* Enhanced User/date/time box */}
                    <div className={styles.headerUserInfoBox}>
                        <div className={styles.userSection}>
                            <div className={styles.userAvatar}>
                                {getUserInitials(user)}
                            </div>
                            <div className={styles.userDetails}>
                                <div className={styles.userName}>
                                    {user ? (user.name || user.username) : "Guest User"}
                                </div>
                                <div className={styles.userStatus}>
                                    {user ? "Online" : "Not logged in"}
                                </div>
                            </div>
                        </div>
                        <div className={styles.dateTimeSection}>
                            <div className={styles.dateDisplay}>
                                <span className={styles.dateIcon}>📅</span>
                                <span className={styles.dateText}>{formatDate(dateTime)}</span>
                            </div>
                            <div className={styles.timeDisplay}>
                                <span className={styles.timeIcon}>🕐</span>
                                <span className={styles.timeText}>{formatTime(dateTime)}</span>
                            </div>
                        </div>
                    </div>
                    
                    <Link to="/" className={styles.headerTitleContainer}>
                        <h3 className={styles.headerTitle}>{t("headerTitle")}</h3>
                    </Link>
                    <nav>
                        <ul className={`${styles.headerNavList} ${menuOpen ? styles.show : ""}`}>
                            <li>
                                <NavLink
                                    to="/"
                                    className={({ isActive }) => (isActive ? styles.headerNavPageLinkActive : styles.headerNavPageLink)}
                                    onClick={() => setMenuOpen(false)}
                                >
                                    {t("chat")}
                                </NavLink>
                            </li>
                            <li>
                                <NavLink
                                    to="/qa"
                                    className={({ isActive }) => (isActive ? styles.headerNavPageLinkActive : styles.headerNavPageLink)}
                                    onClick={() => setMenuOpen(false)}
                                >
                                    {t("qa")}
                                </NavLink>
                            </li>
                            <li>
                                <NavLink
                                    to="/rfp"
                                    className={({ isActive }) => (isActive ? styles.headerNavPageLinkActive : styles.headerNavPageLink)}
                                    onClick={() => setMenuOpen(false)}
                                >
                                    {t("rfp")}
                                </NavLink>
                            </li>
                            <li>
                                <NavLink
                                    to="/score"
                                    className={({ isActive }) => (isActive ? styles.headerNavPageLinkActive : styles.headerNavPageLink)}
                                    onClick={() => setMenuOpen(false)}
                                >
                                    {t("score")}
                                </NavLink>
                            </li>
                        </ul>
                    </nav>
                    <div className={styles.loginMenuContainer}>
                        {useLogin && <LoginButton />}
                        <IconButton
                            iconProps={{ iconName: "GlobalNavButton" }}
                            className={styles.menuToggle}
                            onClick={toggleMenu}
                            ariaLabel={t("labels.toggleMenu")}
                        />
                    </div>
                </div>
            </header>
            {/* Token Usage Bar */}
            <div className={styles.tokenBarContainer}>
                <div className={styles.tokenBarLabel}>
                    Tokens Left:
                    <span
                        className={
                            tokenUsage < 50
                                ? styles.tokenLow
                                : tokenUsage < 80
                                ? styles.tokenMedium
                                : styles.tokenHigh
                        }
                        style={{ marginLeft: 8 }}
                    >
                        {tokenUsage < 50
                            ? "High"
                            : tokenUsage < 80
                            ? "Medium"
                            : "Low"}
                    </span>
                </div>
                <div className={styles.tokenBarOuter}>
                    <div
                        className={styles.tokenBarInner}
                        style={{
                            width: `${100 - tokenUsage}%`,
                            background:
                                tokenUsage < 50
                                    ? "#22c55e"
                                    : tokenUsage < 80
                                    ? "#facc15"
                                    : "#ef4444"
                        }}
                    />
                </div>
            </div>
            <Outlet />
        </div>
    );
};

export default Layout;