-- =============================================================================
-- FinAccruals LedgerFlow — Submission History Key Queries
-- Target:   Supabase (PostgreSQL 15+)
-- Schema:   submission_history table (greenfield — not yet created)
-- Purpose:  Key analytical views for audit log monitoring
-- Author:   Data Analyst task — Community Dreams Foundation
-- =============================================================================

-- -----------------------------------------------------------------------------
-- REFERENCE: Table DDL (agreed schema)
-- Run this first to create the table before testing any query below.
-- -----------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS submission_history (
    submission_id   TEXT        PRIMARY KEY,          -- FA-YYYY-NNNN
    reference_id    TEXT,                             -- QBO/Xero JE ID (null if failed)
    company_id      TEXT        NOT NULL,             -- QuickBooks realmId / Xero org ID
    submitted_by    TEXT        NOT NULL,             -- user email
    journal_date    DATE        NOT NULL,             -- date on the journal entry lines
    submission_date TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    platform        TEXT        NOT NULL,             -- 'QuickBooks' | 'Xero'
    total_debit     NUMERIC(15,2) NOT NULL DEFAULT 0,
    total_credit    NUMERIC(15,2) NOT NULL DEFAULT 0,
    status          TEXT        NOT NULL,             -- see status codes below
    error_message   TEXT                              -- null when status = POSTED or PENDING
    -- Status codes:
    --   POSTED           — successfully posted to platform
    --   PENDING          — submitted, awaiting platform confirmation
    --   VALIDATION_ERROR — failed pre-submission checks (VAL_*)
    --   AUTH_ERROR       — QuickBooks/Xero session expired (AUTH_*)
    --   PLATFORM_ERROR   — platform rejected the entry (PLT_*)
    --   NETWORK_ERROR    — connection lost during submission (NET_*)
);

-- Index to speed up the most common filter patterns
CREATE INDEX IF NOT EXISTS idx_sh_status          ON submission_history (status);
CREATE INDEX IF NOT EXISTS idx_sh_platform        ON submission_history (platform);
CREATE INDEX IF NOT EXISTS idx_sh_submission_date ON submission_history (submission_date);
CREATE INDEX IF NOT EXISTS idx_sh_journal_date    ON submission_history (journal_date);
CREATE INDEX IF NOT EXISTS idx_sh_company         ON submission_history (company_id);
CREATE INDEX IF NOT EXISTS idx_sh_submitted_by    ON submission_history (submitted_by);


-- =============================================================================
-- SECTION 1 — SUBMISSIONS BY STATUS
-- =============================================================================

-- Q1a: Count of submissions grouped by status
WITH status_counts AS (
    SELECT
        status,
        COUNT(*) AS total_count
    FROM submission_history
    GROUP BY status
)
SELECT
    status,
    total_count,
    ROUND(total_count * 100.0 / SUM(total_count) OVER (), 1) AS pct_of_total
FROM status_counts
ORDER BY total_count DESC;

-- Q1b: Status summary for a specific company (single sandbox/org)
SELECT
    status,
    COUNT(*) AS total_count
FROM submission_history
WHERE company_id = :company_id          -- replace :company_id with actual realmId
GROUP BY status
ORDER BY total_count DESC;

-- Q1c: Failed submissions only — with error messages for triage
SELECT
    submission_id,
    submitted_by,
    submission_date,
    platform,
    status,
    error_message
FROM submission_history
WHERE status NOT IN ('POSTED', 'PENDING')
ORDER BY submission_date DESC;


-- =============================================================================
-- SECTION 2 — SUBMISSIONS BY PLATFORM
-- =============================================================================

-- Q2a: Total submissions per platform
SELECT
    platform,
    COUNT(*)                                AS total_submissions,
    COUNT(*) FILTER (WHERE status = 'POSTED')          AS posted,
    COUNT(*) FILTER (WHERE status = 'PENDING')         AS pending,
    COUNT(*) FILTER (WHERE status NOT IN ('POSTED','PENDING')) AS failed,
    ROUND(
        COUNT(*) FILTER (WHERE status = 'POSTED') * 100.0 / NULLIF(COUNT(*), 0),
        1
    )                                       AS success_rate_pct
FROM submission_history
GROUP BY platform
ORDER BY total_submissions DESC;

-- Q2b: Platform breakdown for a specific date range
SELECT
    platform,
    COUNT(*) AS total_submissions,
    SUM(total_debit) FILTER (WHERE status = 'POSTED') AS total_value_posted
FROM submission_history
WHERE submission_date BETWEEN :date_from AND :date_to
GROUP BY platform
ORDER BY platform;


-- =============================================================================
-- SECTION 3 — SUBMISSIONS BY DATE RANGE
-- =============================================================================

-- Q3a: Daily submission count — flexible period using DATE_TRUNC
--      Change 'day' to 'week' or 'month' for coarser granularity
SELECT
    DATE_TRUNC('day', submission_date)::DATE   AS period,
    COUNT(*)                                   AS total_submissions,
    COUNT(*) FILTER (WHERE status = 'POSTED')  AS posted,
    COUNT(*) FILTER (WHERE status NOT IN ('POSTED','PENDING')) AS failed
FROM submission_history
WHERE submission_date >= NOW() - INTERVAL '30 days'
GROUP BY 1
ORDER BY 1;

-- Q3b: Weekly rollup — last 12 weeks
SELECT
    DATE_TRUNC('week', submission_date)::DATE  AS week_starting,
    COUNT(*)                                   AS total_submissions,
    SUM(total_debit) FILTER (WHERE status = 'POSTED') AS value_posted
FROM submission_history
WHERE submission_date >= NOW() - INTERVAL '12 weeks'
GROUP BY 1
ORDER BY 1;

-- Q3c: Monthly rollup — by journal_date (when the entry was dated, not when submitted)
SELECT
    DATE_TRUNC('month', journal_date)::DATE    AS journal_month,
    COUNT(*)                                   AS total_submissions,
    SUM(total_debit) FILTER (WHERE status = 'POSTED') AS value_posted
FROM submission_history
WHERE journal_date >= DATE_TRUNC('year', NOW())   -- year to date
GROUP BY 1
ORDER BY 1;

-- Q3d: Explicit date range filter (submission_date AND journal_date both supported)
SELECT
    submission_id,
    journal_date,
    submission_date,
    platform,
    submitted_by,
    total_debit,
    total_credit,
    status,
    error_message
FROM submission_history
WHERE
    submission_date BETWEEN :sub_date_from AND :sub_date_to
    -- optionally also filter by journal date:
    -- AND journal_date BETWEEN :jnl_date_from AND :jnl_date_to
ORDER BY submission_date DESC;


-- =============================================================================
-- SECTION 4 — ERROR FREQUENCY
-- =============================================================================

-- Q4a: Error frequency by status category
WITH error_counts AS (
    SELECT
        status,
        COUNT(*) AS error_count
    FROM submission_history
    WHERE status NOT IN ('POSTED', 'PENDING')
    GROUP BY status
)
SELECT
    status,
    error_count,
    ROUND(error_count * 100.0 / SUM(error_count) OVER (), 1) AS pct_of_errors
FROM error_counts
ORDER BY error_count DESC;

-- Q4b: Most common error messages (top 10)
SELECT
    error_message,
    status,
    COUNT(*) AS occurrences
FROM submission_history
WHERE error_message IS NOT NULL
GROUP BY error_message, status
ORDER BY occurrences DESC
LIMIT 10;

-- Q4c: Error rate per platform — which platform fails more often?
SELECT
    platform,
    COUNT(*) AS total,
    COUNT(*) FILTER (WHERE status NOT IN ('POSTED','PENDING')) AS errors,
    ROUND(
        COUNT(*) FILTER (WHERE status NOT IN ('POSTED','PENDING')) * 100.0
        / NULLIF(COUNT(*), 0),
        1
    ) AS error_rate_pct
FROM submission_history
GROUP BY platform
ORDER BY error_rate_pct DESC;

-- Q4d: Error trend — errors per day over the last 30 days
SELECT
    DATE_TRUNC('day', submission_date)::DATE AS day,
    COUNT(*) FILTER (WHERE status = 'VALIDATION_ERROR') AS validation_errors,
    COUNT(*) FILTER (WHERE status = 'AUTH_ERROR')       AS auth_errors,
    COUNT(*) FILTER (WHERE status = 'PLATFORM_ERROR')   AS platform_errors,
    COUNT(*) FILTER (WHERE status = 'NETWORK_ERROR')    AS network_errors
FROM submission_history
WHERE
    submission_date >= NOW() - INTERVAL '30 days'
    AND status NOT IN ('POSTED', 'PENDING')
GROUP BY 1
ORDER BY 1;

-- Q4e: Users with the most errors — for targeted support
SELECT
    submitted_by,
    COUNT(*) AS total_submissions,
    COUNT(*) FILTER (WHERE status NOT IN ('POSTED','PENDING')) AS errors,
    ROUND(
        COUNT(*) FILTER (WHERE status NOT IN ('POSTED','PENDING')) * 100.0
        / NULLIF(COUNT(*), 0),
        1
    ) AS error_rate_pct
FROM submission_history
GROUP BY submitted_by
HAVING COUNT(*) > 1
ORDER BY errors DESC;


-- =============================================================================
-- SECTION 5 — BONUS: COMBINED SUMMARY VIEW
-- Useful for a single dashboard query covering all KPIs
-- =============================================================================

SELECT
    COUNT(*)                                                    AS total_submissions,
    COUNT(*) FILTER (WHERE status = 'POSTED')                  AS posted,
    COUNT(*) FILTER (WHERE status = 'PENDING')                 AS pending,
    COUNT(*) FILTER (WHERE status NOT IN ('POSTED','PENDING')) AS failed,
    COALESCE(SUM(total_debit) FILTER (WHERE status = 'POSTED'), 0) AS total_value_posted,
    ROUND(
        COUNT(*) FILTER (WHERE status = 'POSTED') * 100.0
        / NULLIF(COUNT(*), 0),
        1
    )                                                          AS success_rate_pct
FROM submission_history
WHERE submission_date >= DATE_TRUNC('month', NOW());  -- current month; adjust as needed
