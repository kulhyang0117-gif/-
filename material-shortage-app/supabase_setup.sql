-- ================================================================
-- 자재부족현황 앱 Supabase 초기 설정 SQL
-- Supabase 대시보드 → SQL Editor에서 실행하세요.
-- ================================================================

-- ── 1. profiles 테이블 (회원 정보 + 권한) ────────────────────────
CREATE TABLE IF NOT EXISTS profiles (
  id          UUID REFERENCES auth.users(id) ON DELETE CASCADE PRIMARY KEY,
  email       TEXT NOT NULL,
  display_name TEXT NOT NULL,
  status      TEXT NOT NULL DEFAULT 'pending',   -- pending / approved / rejected
  is_admin    BOOLEAN NOT NULL DEFAULT false,
  permissions JSONB NOT NULL DEFAULT '{"read":false,"upload":false,"modify":false}',
  created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
  approved_at TIMESTAMPTZ
);

-- ── 2. upload_logs 테이블 (업로드 이력, 장치간 공유) ─────────────
CREATE TABLE IF NOT EXISTS upload_logs (
  file_type     TEXT PRIMARY KEY,   -- bom / plan / inv / pkg / kit
  uploaded_at   TEXT,               -- Unix ms timestamp (문자열)
  uploader_name TEXT DEFAULT '',
  file_name     TEXT DEFAULT '',    -- 원본 파일명 (단일 파일)
  file_names    TEXT[] DEFAULT '{}' -- 원본 파일명 목록 (키팅 복수 파일)
);

-- 기존 테이블에 컬럼 추가 (이미 생성된 경우)
ALTER TABLE upload_logs ADD COLUMN IF NOT EXISTS file_name  TEXT DEFAULT '';
ALTER TABLE upload_logs ADD COLUMN IF NOT EXISTS file_names TEXT[] DEFAULT '{}';

-- ── 3. Row Level Security 활성화 ────────────────────────────────
ALTER TABLE profiles    ENABLE ROW LEVEL SECURITY;
ALTER TABLE upload_logs ENABLE ROW LEVEL SECURITY;

-- ── 4. profiles RLS 정책 ────────────────────────────────────────
-- 인증된 사용자는 모든 프로필 읽기 가능
CREATE POLICY "profiles_read" ON profiles
  FOR SELECT USING (auth.role() = 'authenticated');

-- 인증된 사용자는 프로필 수정 가능 (관리자 승인/권한 변경용)
CREATE POLICY "profiles_write" ON profiles
  FOR ALL USING (auth.role() = 'authenticated');

-- ── 5. upload_logs RLS 정책 ─────────────────────────────────────
-- 인증된 사용자는 업로드 이력 읽기/쓰기 가능
CREATE POLICY "upload_logs_all" ON upload_logs
  FOR ALL USING (auth.role() = 'authenticated');

-- 통합 대시보드 게스트(비로그인)도 파일 메타/데이터 읽기 허용
CREATE POLICY "upload_logs_read_anon" ON upload_logs
  FOR SELECT USING (true);

-- ── 6. 관리자 계정 초기 설정 ────────────────────────────────────
-- 아래 단계를 따르세요:
--
-- 1) 앱에서 kulhyang0117@gmail.com / jxy0830! 로 "회원가입 신청"
-- 2) Supabase 대시보드 → Authentication → Users 에서 이메일 확인 필요시 confirm
-- 3) 이 SQL로 관리자 권한 부여 (user_id는 위 Users 목록에서 복사):
--
-- UPDATE profiles SET
--   status = 'approved',
--   is_admin = true,
--   permissions = '{"read":true,"upload":true,"modify":true}'::jsonb,
--   approved_at = now()
-- WHERE email = 'kulhyang0117@gmail.com';

-- ── 7. Storage 버킷 정책 (파일 공유) ───────────────────────────
-- Supabase 대시보드 → Storage → ms-files 버킷 → Policies 에서 설정하거나
-- 아래 SQL을 실행하세요. (버킷이 없으면 대시보드에서 먼저 생성)

-- 인증된 회원은 모든 파일 읽기 가능 (링크로 들어와도 이전 업로드 파일 조회)
CREATE POLICY "storage_read_authenticated" ON storage.objects
  FOR SELECT
  USING (bucket_id = 'ms-files' AND auth.role() = 'authenticated');

-- 통합 대시보드 게스트(비로그인)도 파일 읽기 허용
CREATE POLICY "storage_read_anon" ON storage.objects
  FOR SELECT
  USING (bucket_id = 'ms-files');

-- 인증된 회원은 파일 업로드/덮어쓰기 가능 (upsert)
CREATE POLICY "storage_insert_authenticated" ON storage.objects
  FOR INSERT
  WITH CHECK (bucket_id = 'ms-files' AND auth.role() = 'authenticated');

CREATE POLICY "storage_update_authenticated" ON storage.objects
  FOR UPDATE
  USING (bucket_id = 'ms-files' AND auth.role() = 'authenticated');

-- ★ upsert(덮어쓰기) 시 기존 파일 삭제 권한 필요 — 없으면 업로드 실패
CREATE POLICY "storage_delete_authenticated" ON storage.objects
  FOR DELETE
  USING (bucket_id = 'ms-files' AND auth.role() = 'authenticated');

-- ── 이미 정책이 존재하면 아래를 먼저 실행 후 위 CREATE POLICY 재실행 ──
-- DROP POLICY IF EXISTS "storage_read_authenticated"   ON storage.objects;
-- DROP POLICY IF EXISTS "storage_insert_authenticated" ON storage.objects;
-- DROP POLICY IF EXISTS "storage_update_authenticated" ON storage.objects;
-- DROP POLICY IF EXISTS "storage_delete_authenticated" ON storage.objects;

-- ── 8. Supabase Auth 설정 권장사항 ──────────────────────────────
-- Authentication → Settings 에서:
--   - "Enable email confirmations" → OFF (내부 앱이므로 이메일 확인 불필요)
--   - "Allow new users to sign up" → ON

-- ── 9. Kitting Board OK 체크 / 비고 동기화 테이블 ──────────────────
CREATE TABLE IF NOT EXISTS kb_ok_checks (
  date_serial  INTEGER  NOT NULL,
  part_no      TEXT     NOT NULL,
  checked      BOOLEAN  NOT NULL DEFAULT false,
  note         TEXT     NOT NULL DEFAULT '',
  updated_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_by   TEXT     NOT NULL DEFAULT '',
  PRIMARY KEY (date_serial, part_no)
);

ALTER TABLE kb_ok_checks ENABLE ROW LEVEL SECURITY;

CREATE POLICY "kb_ok_checks_all" ON kb_ok_checks
  FOR ALL USING (auth.role() = 'authenticated');

-- 통합 대시보드 게스트도 OK 체크/비고 읽기 허용
CREATE POLICY "kb_ok_checks_read_anon" ON kb_ok_checks
  FOR SELECT USING (true);

-- Realtime 활성화 (Supabase 대시보드 → Database → Replication 에서도 테이블 추가 필요)
ALTER PUBLICATION supabase_realtime ADD TABLE kb_ok_checks;
