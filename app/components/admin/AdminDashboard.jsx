"use client";

import React, { useEffect, useState } from "react";

export default function AdminDashboard({
  styles,
  supabase,
  userEmail,
  userShip,
  onBack,
}) {
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [stats, setStats] = useState({
    inventoryCounts: 0,
    masterItems: 0,
    stationStatuses: 0,
    usageLogs: 0,
  });
  const [recentLogs, setRecentLogs] = useState([]);

  const loadAdminStats = async () => {
    if (!supabase) {
      setMessage("Supabase is not connected.");
      return;
    }

    setLoading(true);
    setMessage("");

    try {
      const getCount = async (tableName) => {
        const { count, error } = await supabase
          .from(tableName)
          .select("id", { count: "exact", head: true });

        if (error) throw error;
        return count || 0;
      };

      const [
        inventoryCounts,
        masterItems,
        stationStatuses,
        usageLogs,
      ] = await Promise.all([
        getCount("inventory_counts"),
        getCount("inventory_master_items"),
        getCount("inventory_station_status"),
        getCount("app_usage_logs"),
      ]);

      setStats({
        inventoryCounts,
        masterItems,
        stationStatuses,
        usageLogs,
      });

      const logsResult = await supabase
        .from("app_usage_logs")
        .select("event_type,module,ship,station,user_email,created_at")
        .order("created_at", { ascending: false })
        .limit(20);

      if (!logsResult.error) {
        setRecentLogs(logsResult.data || []);
      }

      setMessage("Admin dashboard refreshed.");
    } catch (error) {
      setMessage(error?.message || "Could not load admin dashboard.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadAdminStats();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const statCards = [
    {
      title: "Inventory Count Records",
      value: stats.inventoryCounts,
      icon: "📝",
    },
    {
      title: "Master Inventory Items",
      value: stats.masterItems,
      icon: "📦",
    },
    {
      title: "Station Status Records",
      value: stats.stationStatuses,
      icon: "📡",
    },
    {
      title: "Usage Log Events",
      value: stats.usageLogs,
      icon: "📊",
    },
  ];

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />

        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={onBack}>
            ← Modules
          </button>
          <div style={styles.shipBadge}>🛡️ Admin</div>
        </div>
      </header>

      <section style={styles.card}>
        <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 16 }}>
          <div>
            <h2 style={styles.productTitle}>🛡️ Admin Dashboard</h2>
            <p style={{ ...styles.emptyText, margin: 0 }}>
              Signed in as {userEmail || "Unknown"} / Ship {userShip || "N/A"}
            </p>
          </div>

          <button
            style={styles.primaryButton}
            onClick={loadAdminStats}
            disabled={loading}
          >
            {loading ? "Loading..." : "🔄 Refresh Admin Data"}
          </button>
        </div>

        {message && (
          <div style={styles.infoBox}>
            {message}
          </div>
        )}

        <div style={styles.moduleGrid}>
          {statCards.map((item) => (
            <div key={item.title} style={styles.moduleCard}>
              <div style={styles.moduleIcon}>{item.icon}</div>
              <strong>{item.title}</strong>
              <span style={{ fontSize: 28, fontWeight: 900 }}>
                {item.value}
              </span>
            </div>
          ))}
        </div>
      </section>

      <section style={styles.card}>
        <h2 style={styles.productTitle}>📋 Recent Usage Logs</h2>

        {!recentLogs.length && (
          <p style={styles.emptyText}>
            No recent logs loaded, or select permission is not enabled for app_usage_logs.
          </p>
        )}

        <div style={styles.equipmentGrid}>
          {recentLogs.map((log, index) => (
            <div key={`${log.event_type}-${index}`} style={styles.equipmentCard}>
              <div style={styles.recipeName}>{log.event_type || "Event"}</div>
              <div style={styles.recipeMeta}>Module: {log.module || "N/A"}</div>
              <div style={styles.recipeMeta}>Ship: {log.ship || "N/A"}</div>
              <div style={styles.recipeMeta}>Station: {log.station || "N/A"}</div>
              <div style={styles.recipeMeta}>User: {log.user_email || "N/A"}</div>
              <div style={styles.recipeMeta}>
                Time: {log.created_at ? new Date(log.created_at).toLocaleString() : "N/A"}
              </div>
            </div>
          ))}
        </div>
      </section>
    </main>
  );
}
