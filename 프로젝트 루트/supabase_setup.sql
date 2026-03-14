-- ================================================================
-- 자재부족현황 앱 Supabase 초기 설정 SQL
-- Supabase 대시보드 → SQL Editor에서 실행하세요.
-- ================================================================

-- ── 1. profiles 테이블 ────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS profiles (
  id          UUID REFERENCES auth.users(id) ON DELETE CASCADE PRIMARY KEY,
  email       TEXT NOT NULL,
  display_name TEXT NOT NULL,
  status      TEXT NOT NULL DEFAULT 'pending',
  is_admin    BOOLEAN NOT NULL DEFAULT false,
  permissions JSONB NOT NULL DEFAULT '{"read":false,"upload":false,"modify":false}',
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
  approved_at TIMESTAMPTZ
);

-- ── 2. upload_logs 테이블 (업로드 이력 + 키팅 파일 목록) ───────────
CREATE TABLE IF NOT EXISTS upload_logs (
  file_type     TEXT PRIMARY KEY,   -- bom / plan / inv / kit
  uploaded_at   TEXT,               -- Unix ms timestamp
  uploader_name TEXT DEFAULT '',
  file_names    JSONB DEFAULT '[]'  -- 키팅 파일 목록
);

-- ── 3. Row Level Security ─────────────────────────────────────────
ALTER TABLE profiles    ENABLE ROW LEVEL SECURITY;
ALTER TABLE upload_logs ENABLE ROW LEVEL SECURITY;

-- ── 4. RLS 정책 ───────────────────────────────────────────────────
CREATE POLICY "profiles_all"     ON profiles    FOR ALL USING (auth.role() = 'authenticated');
CREATE POLICY "upload_logs_all"  ON upload_logs FOR ALL USING (auth.role() = 'authenticated');

-- ── 5. Storage RLS (Storage 버킷 생성 후 실행) ───────────────────
CREATE POLICY "ms_files_auth" ON storage.objects
  FOR ALL TO authenticated
  USING (bucket_id = 'ms-files')
  WITH CHECK (bucket_id = 'ms-files');

-- ── 6. 관리자 권한 부여 (회원가입 후 실행) ───────────────────────
-- UPDATE profiles SET
--   status = 'approved', is_admin = true,
--   permissions = '{"read":true,"upload":true,"modify":true}'::jsonb,
--   approved_at = now()
-- WHERE email = 'kulhyang0117@gmail.com';
