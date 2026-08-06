-- PostgreSQL bootstrap. The application also runs idempotent startup migrations.
CREATE TABLE IF NOT EXISTS schema_migrations(
  version TEXT PRIMARY KEY,
  applied_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
