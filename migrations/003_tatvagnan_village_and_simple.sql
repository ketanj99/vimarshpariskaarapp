-- Tatvagnan simple data: one row per ADD/REMOVE (group, language, action_type); summary/group PDF use this
CREATE TABLE IF NOT EXISTS vimars.tatvagnan_simple_data (
  id SERIAL PRIMARY KEY,
  pushp_no TEXT NOT NULL,
  "group" TEXT NOT NULL,
  language TEXT NOT NULL,
  action_type TEXT NOT NULL CHECK (action_type IN ('ADD', 'REMOVE')),
  created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_tatvagnan_simple_pushp ON vimars.tatvagnan_simple_data(pushp_no);
CREATE INDEX IF NOT EXISTS idx_tatvagnan_simple_group ON vimars.tatvagnan_simple_data("group");

-- Extend for group+village wise table with names (no new table)
ALTER TABLE vimars.tatvagnan_simple_data ADD COLUMN IF NOT EXISTS village TEXT;
ALTER TABLE vimars.tatvagnan_simple_data ADD COLUMN IF NOT EXISTS name TEXT;
ALTER TABLE vimars.tatvagnan_simple_data ADD COLUMN IF NOT EXISTS member TEXT;
ALTER TABLE vimars.tatvagnan_simple_data ADD COLUMN IF NOT EXISTS mobile_number TEXT;
ALTER TABLE vimars.tatvagnan_simple_data ADD COLUMN IF NOT EXISTS pin_code TEXT;
ALTER TABLE vimars.tatvagnan_simple_data ADD COLUMN IF NOT EXISTS address_line1 TEXT;
ALTER TABLE vimars.tatvagnan_simple_data ADD COLUMN IF NOT EXISTS address_line2 TEXT;
CREATE INDEX IF NOT EXISTS idx_tatvagnan_simple_village ON vimars.tatvagnan_simple_data(village);

COMMENT ON TABLE vimars.tatvagnan_simple_data IS 'One row per ADD/REMOVE; group+language for PDF; village+name for group+village wise table';
