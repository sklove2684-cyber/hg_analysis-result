PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS analysis_batches (
    batch_id TEXT PRIMARY KEY,
    batch_code TEXT NOT NULL UNIQUE,
    analysis_type TEXT NOT NULL,
    analysis_no_start INTEGER NOT NULL,
    analysis_no_end INTEGER NOT NULL,
    parser_name TEXT NOT NULL,
    parser_version TEXT NOT NULL,
    parser_layout_id TEXT NOT NULL,
    extracted_at TEXT NOT NULL,
    warning_count INTEGER NOT NULL DEFAULT 0 CHECK (warning_count >= 0),
    review_status TEXT NOT NULL,
    workplace TEXT,
    year INTEGER,
    period TEXT,
    device_id TEXT,
    analyst TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    CHECK (analysis_no_start <= analysis_no_end)
);

CREATE TABLE IF NOT EXISTS source_files (
    source_file_id TEXT PRIMARY KEY,
    batch_id TEXT NOT NULL UNIQUE REFERENCES analysis_batches(batch_id) ON DELETE CASCADE,
    original_name TEXT NOT NULL,
    full_path TEXT NOT NULL,
    file_hash TEXT NOT NULL UNIQUE CHECK (length(file_hash) = 64),
    file_size INTEGER NOT NULL CHECK (file_size >= 0),
    page_count INTEGER NOT NULL CHECK (page_count >= 1),
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS samples (
    sample_id TEXT PRIMARY KEY,
    batch_id TEXT NOT NULL REFERENCES analysis_batches(batch_id) ON DELETE CASCADE,
    page_no INTEGER NOT NULL CHECK (page_no >= 1),
    sample_name_raw TEXT NOT NULL,
    sample_name_normalized TEXT NOT NULL,
    data_filename TEXT,
    method_filename TEXT,
    batch_filename TEXT,
    acquired_at TEXT,
    sample_type TEXT NOT NULL,
    concentration_level TEXT,
    replicate_no INTEGER,
    worker_match_key TEXT,
    is_blank INTEGER NOT NULL DEFAULT 0 CHECK (is_blank IN (0, 1)),
    total_area INTEGER,
    created_at TEXT NOT NULL,
    UNIQUE (batch_id, page_no)
);

CREATE TABLE IF NOT EXISTS peaks (
    peak_id TEXT PRIMARY KEY,
    sample_id TEXT NOT NULL REFERENCES samples(sample_id) ON DELETE CASCADE,
    peak_no INTEGER NOT NULL CHECK (peak_no >= 1),
    retention_time TEXT NOT NULL,
    area_raw INTEGER NOT NULL CHECK (area_raw >= 0),
    height INTEGER CHECK (height IS NULL OR height >= 0),
    material_raw TEXT,
    material_standard TEXT,
    peak_group_no INTEGER,
    include_for_excel INTEGER NOT NULL DEFAULT 1 CHECK (include_for_excel IN (0, 1)),
    exclude_reason TEXT,
    source_page INTEGER NOT NULL CHECK (source_page >= 0),
    created_at TEXT NOT NULL,
    UNIQUE (sample_id, peak_no)
);

CREATE TABLE IF NOT EXISTS peak_corrections (
    correction_id TEXT PRIMARY KEY,
    peak_id TEXT NOT NULL REFERENCES peaks(peak_id) ON DELETE CASCADE,
    area_before INTEGER NOT NULL CHECK (area_before >= 0),
    area_after INTEGER NOT NULL CHECK (area_after >= 0),
    reason TEXT NOT NULL,
    corrected_at TEXT NOT NULL,
    device_id TEXT NOT NULL,
    revision_no INTEGER NOT NULL CHECK (revision_no >= 1),
    UNIQUE (peak_id, revision_no)
);

CREATE TABLE IF NOT EXISTS export_jobs (
    export_job_id TEXT PRIMARY KEY,
    batch_id TEXT NOT NULL REFERENCES analysis_batches(batch_id) ON DELETE CASCADE,
    template_path TEXT NOT NULL,
    output_path TEXT NOT NULL,
    std_method TEXT NOT NULL CHECK (std_method IN ('A', 'B')),
    device_id TEXT NOT NULL,
    status TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_batches_search
    ON analysis_batches(workplace, year, period, analysis_type);
CREATE INDEX IF NOT EXISTS idx_samples_batch_name
    ON samples(batch_id, sample_name_normalized);
CREATE INDEX IF NOT EXISTS idx_peaks_sample_material
    ON peaks(sample_id, material_standard);
CREATE INDEX IF NOT EXISTS idx_corrections_peak_revision
    ON peak_corrections(peak_id, revision_no DESC);
CREATE INDEX IF NOT EXISTS idx_exports_batch
    ON export_jobs(batch_id);

CREATE TRIGGER IF NOT EXISTS prevent_peak_area_raw_update
BEFORE UPDATE OF area_raw ON peaks
FOR EACH ROW
WHEN NEW.area_raw <> OLD.area_raw
BEGIN
    SELECT RAISE(ABORT, 'peaks.area_raw is immutable; add a peak_correction');
END;
