-- Step 1: Add access_code to teams table
ALTER TABLE teams
ADD COLUMN access_code VARCHAR(10) UNIQUE;

-- Step 2: Update existing teams with random access codes
UPDATE teams
SET access_code = 
  UPPER(
    SUBSTRING(
      MD5(RANDOM()::TEXT || id::TEXT) 
      FROM 1 FOR 8
    )
  )
WHERE access_code IS NULL;

-- Step 3: Make access_code NOT NULL after populating existing teams
ALTER TABLE teams
ALTER COLUMN access_code SET NOT NULL;

-- Step 4: Create join_requests table
CREATE TABLE join_requests (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id UUID REFERENCES users(id) ON DELETE CASCADE NOT NULL,
  team_id UUID REFERENCES teams(id) ON DELETE CASCADE NOT NULL,
  request_date TIMESTAMP WITH TIME ZONE DEFAULT NOW() NOT NULL,
  status VARCHAR(20) DEFAULT 'pending' NOT NULL, -- 'pending', 'approved', 'rejected'
  resolved_date TIMESTAMP WITH TIME ZONE,
  resolved_by UUID REFERENCES users(id) ON DELETE SET NULL,
  
  -- Ensure a user can only have one active request per team
  CONSTRAINT unique_user_team_request UNIQUE (user_id, team_id, status),
  CONSTRAINT valid_status CHECK (status IN ('pending', 'approved', 'rejected'))
);

-- Step 5: Add index for faster lookups
CREATE INDEX join_requests_user_id_idx ON join_requests(user_id);
CREATE INDEX join_requests_team_id_idx ON join_requests(team_id);
CREATE INDEX join_requests_status_idx ON join_requests(status);

-- Step 6: Create RLS policies for join_requests table

-- Only allow users to see their own requests and admins to see their team's requests
CREATE POLICY join_requests_select_policy ON join_requests
  FOR SELECT USING (
    auth.uid() = user_id OR 
    EXISTS (
      SELECT 1 FROM users 
      WHERE users.id = auth.uid() 
      AND users.team_id = join_requests.team_id 
      AND users.role = 'admin'
    )
  );

-- Only allow users to create their own requests
CREATE POLICY join_requests_insert_policy ON join_requests
  FOR INSERT WITH CHECK (auth.uid() = user_id);

-- Only allow admins to update requests for their team
CREATE POLICY join_requests_update_policy ON join_requests
  FOR UPDATE USING (
    EXISTS (
      SELECT 1 FROM users 
      WHERE users.id = auth.uid() 
      AND users.team_id = join_requests.team_id 
      AND users.role = 'admin'
    )
  );

-- Enable RLS for join_requests
ALTER TABLE join_requests ENABLE ROW LEVEL SECURITY; 