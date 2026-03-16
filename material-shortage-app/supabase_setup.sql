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
  file_type     TEXT PRIMARY KEY,   -- bom / plan / inv / kit
  uploaded_at   TEXT,               -- Unix ms timestamp (문자열)
  uploader_name TEXT DEFAULT ''
);

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

-- ── 7. Supabase Auth 설정 권장사항 ──────────────────────────────
-- Authentication → Settings 에서:
--   - "Enable email confirmations" → OFF (내부 앱이므로 이메일 확인 불필요)
--   - "Allow new users to sign up" → ON
