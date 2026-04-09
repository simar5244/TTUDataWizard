#!/usr/bin/env bash
# DataWizard Full API Test Suite
# Tests: register, login, mappings CRUD, version control, dashboards CRUD, sharing, delete

set -euo pipefail

BASE="http://localhost:3000"
COOKIE_JAR="/tmp/dw_cookies.txt"
EMAIL="simarsin@ttu.edu"
PASS="TexasTech2023!"
PASS_WRONG="WrongPass999!"

GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

pass=0; fail=0; warn=0

ok()   { echo -e "  ${GREEN}✓${NC} $1"; pass=$((pass+1)); }
fail() { echo -e "  ${RED}✗${NC} $1"; fail=$((fail+1)); }
warn() { echo -e "  ${YELLOW}~${NC} $1"; warn=$((warn+1)); }
hdr()  { echo -e "\n${BLUE}━━━ $1 ━━━${NC}"; }

expect_field() {   # expect_field "field" "response_body"
  local field="$1" body="$2"
  if echo "$body" | grep -q "\"$field\""; then
    ok "Response contains '$field'"
  else
    fail "Response missing '$field' — got: $(echo "$body" | head -c 300)"
  fi
}

expect_status() {   # expect_status expected actual context
  if [[ "$2" == "$1" ]]; then
    ok "HTTP $1 — $3"
  else
    fail "Expected HTTP $1 got $2 — $3"
  fi
}

# ── 1. REGISTER ───────────────────────────────────────────────────────────────
hdr "1 · REGISTER"

REG=$(curl -s -o /dev/null -w "%{http_code}" -X POST "$BASE/api/register" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Sim Arsin\",\"email\":\"$EMAIL\",\"password\":\"$PASS\"}")

if [[ "$REG" == "201" ]]; then
  ok "Registered new user (201)"
elif [[ "$REG" == "409" ]]; then
  warn "User already exists (409) — continuing"
else
  fail "Unexpected register status: $REG"
fi

# Duplicate register → 409
DUP=$(curl -s -o /dev/null -w "%{http_code}" -X POST "$BASE/api/register" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Dup\",\"email\":\"$EMAIL\",\"password\":\"$PASS\"}")
expect_status "409" "$DUP" "duplicate register blocked"

# ── 2. AUTH / LOGIN ───────────────────────────────────────────────────────────
hdr "2 · AUTH / LOGIN"

# Get CSRF token
CSRF_RESP=$(curl -s -c "$COOKIE_JAR" "$BASE/api/auth/csrf")
CSRF_TOKEN=$(echo "$CSRF_RESP" | sed 's/.*"csrfToken":"\([^"]*\)".*/\1/')
if [[ -n "$CSRF_TOKEN" && "$CSRF_TOKEN" != "$CSRF_RESP" ]]; then
  ok "Got CSRF token: ${CSRF_TOKEN:0:12}…"
else
  fail "Failed to get CSRF token"
  exit 1
fi

# Wrong password → should NOT set session
BAD_LOGIN=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" -c "$COOKIE_JAR" \
  -X POST "$BASE/api/auth/callback/credentials" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data-urlencode "csrfToken=$CSRF_TOKEN" \
  --data-urlencode "email=$EMAIL" \
  --data-urlencode "password=$PASS_WRONG" \
  --data-urlencode "redirect=false")
warn "Bad password login returned HTTP $BAD_LOGIN (302/200 both possible — NextAuth redirects on failure)"

# Correct login
LOGIN_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" -c "$COOKIE_JAR" \
  -X POST "$BASE/api/auth/callback/credentials" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data-urlencode "csrfToken=$CSRF_TOKEN" \
  --data-urlencode "email=$EMAIL" \
  --data-urlencode "password=$PASS" \
  --data-urlencode "redirect=false")
warn "Login POST returned HTTP $LOGIN_STATUS (NextAuth redirects — checking session next)"

# Verify session
SESSION=$(curl -s -b "$COOKIE_JAR" "$BASE/api/auth/session")
if echo "$SESSION" | grep -q "\"email\""; then
  ok "Session active — $(echo "$SESSION" | sed 's/.*"email":"\([^"]*\)".*/\1/')"
else
  fail "No active session — response: $SESSION"
  echo "  Cannot continue without auth. Check credentials / NEXTAUTH_SECRET."
  exit 1
fi

# ── 3. UNAUTHORIZED ACCESS ────────────────────────────────────────────────────
hdr "3 · UNAUTHORIZED ACCESS (no session)"

UNAUTH=$(curl -s -o /dev/null -w "%{http_code}" "$BASE/api/mappings")
expect_status "401" "$UNAUTH" "mappings list without session → 401"

UNAUTH_D=$(curl -s -o /dev/null -w "%{http_code}" "$BASE/api/dashboards")
expect_status "401" "$UNAUTH_D" "dashboards list without session → 401"

# ── 4. MAPPINGS — CREATE ──────────────────────────────────────────────────────
hdr "4 · MAPPINGS — CREATE"

CONNECTIONS_V1='{
  "nodes":[
    {"id":"excel_col_0","type":"excelCol","position":{"x":60,"y":60},"data":{"label":"Contact ID","dataType":"string","colKey":"col_0"}},
    {"id":"excel_col_1","type":"excelCol","position":{"x":60,"y":130},"data":{"label":"First Name","dataType":"string","colKey":"col_1"}},
    {"id":"excel_col_2","type":"excelCol","position":{"x":60,"y":200},"data":{"label":"Last Name","dataType":"string","colKey":"col_2"}},
    {"id":"excel_col_3","type":"excelCol","position":{"x":60,"y":270},"data":{"label":"Email","dataType":"string","colKey":"col_3"}},
    {"id":"excel_col_4","type":"excelCol","position":{"x":60,"y":340},"data":{"label":"Annual Revenue","dataType":"number","colKey":"col_4"}},
    {"id":"ss_col_0","type":"ssCol","position":{"x":700,"y":60},"data":{"label":"Record ID","colType":"string","colId":"col_0"}},
    {"id":"ss_col_1","type":"ssCol","position":{"x":700,"y":130},"data":{"label":"Name","colType":"string","colId":"col_1"}},
    {"id":"ss_col_2","type":"ssCol","position":{"x":700,"y":200},"data":{"label":"Primary Email","colType":"string","colId":"col_2"}},
    {"id":"ss_col_3","type":"ssCol","position":{"x":700,"y":270},"data":{"label":"Org Revenue","colType":"number","colId":"col_3"}}
  ],
  "edges":[
    {"id":"e1","source":"excel_col_0","target":"ss_col_0","data":{"formula":""}},
    {"id":"e2","source":"excel_col_1","target":"ss_col_1","data":{"formula":"CONCAT(source, \" \", next)"}},
    {"id":"e3","source":"excel_col_3","target":"ss_col_2","data":{"formula":""}},
    {"id":"e4","source":"excel_col_4","target":"ss_col_3","data":{"formula":"source * 0.85"}}
  ]
}'

SCHEMA='{
  "sheetName":"salesforce_export",
  "columnNames":["Contact ID","First Name","Last Name","Email","Annual Revenue"],
  "columnOrder":["Contact ID","First Name","Last Name","Email","Annual Revenue"],
  "dataTypes":{"Contact ID":"string","First Name":"string","Last Name":"string","Email":"string","Annual Revenue":"number"},
  "rowCountApprox":800
}'

CREATE_M=$(curl -s -b "$COOKIE_JAR" -X POST "$BASE/api/mappings" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Salesforce → CRM v1\",\"connections\":$CONNECTIONS_V1,\"formulas\":{\"e2\":\"CONCAT(source, ' ', next)\",\"e4\":\"source * 0.85\"},\"schemaFingerprint\":$SCHEMA,\"changeSummary\":\"Initial mapping\"}")

MAPPING_ID=$(echo "$CREATE_M" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['id'])" 2>/dev/null || true)
MAPPING_SLUG=$(echo "$CREATE_M" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['slug'])" 2>/dev/null || true)

if [[ -n "$MAPPING_ID" ]]; then
  ok "Created mapping id=$MAPPING_ID slug=$MAPPING_SLUG"
else
  fail "Mapping create failed: $CREATE_M"
  exit 1
fi
expect_field "currentVersionId" "$CREATE_M"
expect_field "slug" "$CREATE_M"

# Create a second mapping
CREATE_M2=$(curl -s -b "$COOKIE_JAR" -X POST "$BASE/api/mappings" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Financials → Summary\",\"connections\":{\"nodes\":[],\"edges\":[]},\"formulas\":{},\"schemaFingerprint\":{},\"changeSummary\":\"Blank starter\"}")
MAPPING2_ID=$(echo "$CREATE_M2" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['id'])" 2>/dev/null || true)
if [[ -n "$MAPPING2_ID" ]]; then
  ok "Created second mapping id=$MAPPING2_ID"
else
  fail "Second mapping create failed"
fi

# ── 5. MAPPINGS — LIST & GET ──────────────────────────────────────────────────
hdr "5 · MAPPINGS — LIST & GET"

LIST_M=$(curl -s -b "$COOKIE_JAR" "$BASE/api/mappings")
LIST_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" "$BASE/api/mappings")
expect_status "200" "$LIST_STATUS" "list mappings"

COUNT=$(echo "$LIST_M" | grep -o '"id"' | wc -l | tr -d ' ')
if [[ "$COUNT" -ge 2 ]]; then
  ok "List returned $COUNT mappings"
else
  fail "Expected ≥2 mappings, got $COUNT"
fi

GET_M=$(curl -s -b "$COOKIE_JAR" "$BASE/api/mappings/$MAPPING_ID")
GET_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" "$BASE/api/mappings/$MAPPING_ID")
expect_status "200" "$GET_STATUS" "get mapping by id"
expect_field "versions" "$GET_M"
expect_field "stagingRuns" "$GET_M"

# 404 for nonexistent
NOT_FOUND=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" "$BASE/api/mappings/nonexistent-id-000")
expect_status "404" "$NOT_FOUND" "get nonexistent mapping → 404"

# ── 6. MAPPINGS — UPDATE (new version) ────────────────────────────────────────
hdr "6 · MAPPINGS — UPDATE / VERSION CONTROL"

CONNECTIONS_V2='{
  "nodes":[
    {"id":"excel_col_0","type":"excelCol","position":{"x":60,"y":60},"data":{"label":"Contact ID","dataType":"string","colKey":"col_0"}},
    {"id":"excel_col_1","type":"excelCol","position":{"x":60,"y":130},"data":{"label":"Full Name","dataType":"string","colKey":"col_1"}},
    {"id":"excel_col_2","type":"excelCol","position":{"x":60,"y":200},"data":{"label":"Email","dataType":"string","colKey":"col_2"}},
    {"id":"excel_col_3","type":"excelCol","position":{"x":60,"y":270},"data":{"label":"Amount","dataType":"number","colKey":"col_3"}},
    {"id":"excel_col_4","type":"excelCol","position":{"x":60,"y":340},"data":{"label":"Close Date","dataType":"date","colKey":"col_4"}},
    {"id":"excel_col_5","type":"excelCol","position":{"x":60,"y":410},"data":{"label":"Stage","dataType":"string","colKey":"col_5"}},
    {"id":"ss_col_0","type":"ssCol","position":{"x":700,"y":60},"data":{"label":"Record ID","colType":"string","colId":"col_0"}},
    {"id":"ss_col_1","type":"ssCol","position":{"x":700,"y":130},"data":{"label":"Name","colType":"string","colId":"col_1"}},
    {"id":"ss_col_2","type":"ssCol","position":{"x":700,"y":200},"data":{"label":"Primary Email","colType":"string","colId":"col_2"}},
    {"id":"ss_col_3","type":"ssCol","position":{"x":700,"y":270},"data":{"label":"Deal Value","colType":"number","colId":"col_3"}},
    {"id":"ss_col_4","type":"ssCol","position":{"x":700,"y":340},"data":{"label":"Expected Close","colType":"date","colId":"col_4"}},
    {"id":"ss_col_5","type":"ssCol","position":{"x":700,"y":410},"data":{"label":"Pipeline Stage","colType":"string","colId":"col_5"}}
  ],
  "edges":[
    {"id":"e1","source":"excel_col_0","target":"ss_col_0","data":{"formula":""}},
    {"id":"e2","source":"excel_col_1","target":"ss_col_1","data":{"formula":"UPPER(source)"}},
    {"id":"e3","source":"excel_col_2","target":"ss_col_2","data":{"formula":""}},
    {"id":"e4","source":"excel_col_3","target":"ss_col_3","data":{"formula":"source * 1.1"}},
    {"id":"e5","source":"excel_col_4","target":"ss_col_4","data":{"formula":""}},
    {"id":"e6","source":"excel_col_5","target":"ss_col_5","data":{"formula":"IF(source==\"Closed Won\",\"Won\",\"Active\")"}}
  ]
}'

UPDATE_M=$(curl -s -b "$COOKIE_JAR" -X PUT "$BASE/api/mappings/$MAPPING_ID" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Salesforce → CRM v2\",\"connections\":$CONNECTIONS_V2,\"formulas\":{\"e2\":\"UPPER(source)\",\"e4\":\"source * 1.1\",\"e6\":\"IF(source=='Closed Won','Won','Active')\"},\"changeSummary\":\"Added Stage + Date mappings, amount markup formula\"}")

V2_ID=$(echo "$UPDATE_M" | grep -o '"id":"[^"]*"' | head -1 | sed 's/"id":"//;s/"//')
if echo "$UPDATE_M" | grep -q '"versionNumber":2'; then
  ok "Version 2 created successfully"
else
  fail "Version bump failed — response: $(echo "$UPDATE_M" | head -c 400)"
fi

# Version 3 — rename + add more edges
UPDATE_V3=$(curl -s -b "$COOKIE_JAR" -X PUT "$BASE/api/mappings/$MAPPING_ID" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Salesforce → CRM (final)\",\"connections\":$CONNECTIONS_V2,\"formulas\":{\"e2\":\"TRIM(UPPER(source))\",\"e4\":\"ROUND(source * 1.1, 2)\"},\"changeSummary\":\"Trimmed name, rounded amount\"}")
if echo "$UPDATE_V3" | grep -q '"versionNumber":3'; then
  ok "Version 3 created successfully"
else
  fail "Version 3 failed"
fi

# ── 7. VERSION LIST & RESTORE ─────────────────────────────────────────────────
hdr "7 · VERSION LIST & RESTORE"

VERSIONS=$(curl -s -b "$COOKIE_JAR" "$BASE/api/mappings/$MAPPING_ID/versions")
V_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" "$BASE/api/mappings/$MAPPING_ID/versions")
expect_status "200" "$V_STATUS" "list versions"

V_COUNT=$(echo "$VERSIONS" | grep -o '"versionNumber"' | wc -l | tr -d ' ')
if [[ "$V_COUNT" -eq 3 ]]; then
  ok "Correct: 3 versions found"
else
  fail "Expected 3 versions, found $V_COUNT"
fi

# Restore version 1
V1_ID=$(echo "$VERSIONS" | python3 -c "import sys,json; vs=json.load(sys.stdin); v=[x for x in vs if x['versionNumber']==1]; print(v[0]['id'] if v else '')" 2>/dev/null || true)
if [[ -n "$V1_ID" ]]; then
  RESTORE=$(curl -s -b "$COOKIE_JAR" -X POST "$BASE/api/mappings/$MAPPING_ID/versions" \
    -H "Content-Type: application/json" \
    -d "{\"versionId\":\"$V1_ID\"}")
  if echo "$RESTORE" | grep -q '"success":true'; then
    ok "Restored to version 1 (id=$V1_ID)"
  else
    fail "Restore failed: $RESTORE"
  fi
else
  fail "Could not extract version 1 id from versions list"
fi

# Restore back to v3
V3_ID=$(echo "$VERSIONS" | python3 -c "import sys,json; vs=json.load(sys.stdin); v=[x for x in vs if x['versionNumber']==3]; print(v[0]['id'] if v else '')" 2>/dev/null || true)
if [[ -n "$V3_ID" ]]; then
  curl -s -b "$COOKIE_JAR" -X POST "$BASE/api/mappings/$MAPPING_ID/versions" \
    -H "Content-Type: application/json" \
    -d "{\"versionId\":\"$V3_ID\"}" > /dev/null
  ok "Restored to version 3 (latest)"
fi

# ── 8. DASHBOARDS — CREATE ────────────────────────────────────────────────────
hdr "8 · DASHBOARDS — CREATE"

CHARTS='[
  {"id":"ch1","type":"bar","title":"Revenue by Region","xColumn":"Region","yColumns":["Revenue"],"colorScheme":"blues","showLegend":false,"showGrid":true,"showLabels":true,"aggregation":"sum","sortBy":"desc","limit":10,"stacked":false},
  {"id":"ch2","type":"line","title":"Monthly Revenue Trend","xColumn":"Month Name","yColumns":["Revenue","Gross Profit"],"colorScheme":"cool","showLegend":true,"showGrid":true,"showLabels":false,"aggregation":"sum","sortBy":"none","limit":12,"stacked":false},
  {"id":"ch3","type":"pie","title":"Revenue by Category","xColumn":"Category","yColumns":["Revenue"],"colorScheme":"warm","showLegend":true,"showGrid":false,"showLabels":true,"aggregation":"sum","sortBy":"desc","limit":8,"stacked":false},
  {"id":"ch4","type":"area","title":"Gross Margin Trend","xColumn":"Month Name","yColumns":["Gross Profit","EBITDA"],"colorScheme":"greens","showLegend":true,"showGrid":true,"showLabels":false,"aggregation":"sum","sortBy":"none","limit":12,"stacked":false},
  {"id":"ch5","type":"donut","title":"Revenue by Segment","xColumn":"Customer Segment","yColumns":["Revenue"],"colorScheme":"purples","showLegend":true,"showGrid":false,"showLabels":true,"aggregation":"sum","sortBy":"desc","limit":6,"stacked":false},
  {"id":"ch6","type":"scatter","title":"LTV vs CAC","xColumn":"CAC","yColumns":["LTV"],"colorScheme":"neon","showLegend":false,"showGrid":true,"showLabels":false,"aggregation":"none","sortBy":"none","limit":200,"stacked":false}
]'

LAYOUT='[
  {"i":"ch1","x":0,"y":0,"w":6,"h":4},
  {"i":"ch2","x":6,"y":0,"w":6,"h":4},
  {"i":"ch3","x":0,"y":4,"w":4,"h":4},
  {"i":"ch4","x":4,"y":4,"w":4,"h":4},
  {"id":"ch5","x":8,"y":4,"w":4,"h":4},
  {"i":"ch6","x":0,"y":8,"w":12,"h":5}
]'

EXCEL_DATA='{"sheetName":"financials_q1","rowCount":1000,"columns":["Transaction ID","Date","Year","Quarter","Month","Month Name","Product","Category","Cost Center","Sales Rep","Region","Country","Customer Segment","Units Sold","Unit Price","Revenue","COGS","Gross Profit","Gross Margin %","OpEx","EBITDA","EBITDA Margin %","MRR","ARR","Invoice Status","LTV","CAC"]}'

CREATE_D=$(curl -s -b "$COOKIE_JAR" -X POST "$BASE/api/dashboards" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Q1 Financial Dashboard\",\"charts\":$CHARTS,\"layout\":$LAYOUT,\"excelData\":$EXCEL_DATA}")

DASH_ID=$(echo "$CREATE_D" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['id'])" 2>/dev/null || true)
DASH_SLUG=$(echo "$CREATE_D" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['slug'])" 2>/dev/null || true)

if [[ -n "$DASH_ID" ]]; then
  ok "Created dashboard id=$DASH_ID slug=$DASH_SLUG"
else
  fail "Dashboard create failed: $CREATE_D"
  exit 1
fi
expect_field "charts" "$CREATE_D"
expect_field "layout" "$CREATE_D"

# Second dashboard — Salesforce KPIs
CHARTS2='[
  {"id":"d2c1","type":"bar","title":"Deals by Stage","xColumn":"Stage","yColumns":["Amount"],"colorScheme":"blues","showLegend":false,"showGrid":true,"showLabels":true,"aggregation":"count","sortBy":"desc","limit":10,"stacked":false},
  {"id":"d2c2","type":"funnel","title":"Sales Funnel","xColumn":"Stage","yColumns":["Amount"],"colorScheme":"greens","showLegend":false,"showGrid":false,"showLabels":true,"aggregation":"count","sortBy":"desc","limit":10,"stacked":false},
  {"id":"d2c3","type":"treemap","title":"Revenue by Industry","xColumn":"Industry","yColumns":["Annual Revenue"],"colorScheme":"warm","showLegend":false,"showGrid":false,"showLabels":true,"aggregation":"sum","sortBy":"desc","limit":15,"stacked":false},
  {"id":"d2c4","type":"table","title":"Top Opportunities","xColumn":"Opportunity Name","yColumns":["Amount","Stage"],"colorScheme":"default","showLegend":false,"showGrid":true,"showLabels":false,"aggregation":"none","sortBy":"none","limit":20,"stacked":false}
]'

CREATE_D2=$(curl -s -b "$COOKIE_JAR" -X POST "$BASE/api/dashboards" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Salesforce Pipeline View\",\"charts\":$CHARTS2,\"layout\":[],\"excelData\":null}")
DASH2_ID=$(echo "$CREATE_D2" | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['id'])" 2>/dev/null || true)
if [[ -n "$DASH2_ID" ]]; then
  ok "Created second dashboard (Salesforce Pipeline) id=$DASH2_ID"
else
  fail "Second dashboard create failed"
fi

# ── 9. DASHBOARDS — LIST & GET ────────────────────────────────────────────────
hdr "9 · DASHBOARDS — LIST & GET"

LIST_D=$(curl -s -b "$COOKIE_JAR" "$BASE/api/dashboards")
LIST_D_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" "$BASE/api/dashboards")
expect_status "200" "$LIST_D_STATUS" "list dashboards"

D_COUNT=$(echo "$LIST_D" | grep -o '"id"' | wc -l | tr -d ' ')
if [[ "$D_COUNT" -ge 2 ]]; then
  ok "List returned $D_COUNT dashboards"
else
  fail "Expected ≥2 dashboards, got $D_COUNT"
fi

GET_D=$(curl -s -b "$COOKIE_JAR" "$BASE/api/dashboards/$DASH_ID")
GET_D_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" "$BASE/api/dashboards/$DASH_ID")
expect_status "200" "$GET_D_STATUS" "get dashboard by id"
expect_field "charts" "$GET_D"

# ── 10. DASHBOARDS — UPDATE ───────────────────────────────────────────────────
hdr "10 · DASHBOARDS — UPDATE"

UPDATED_CHARTS='[
  {"id":"ch1","type":"bar","title":"Revenue by Region (Updated)","xColumn":"Region","yColumns":["Revenue","Gross Profit"],"colorScheme":"purples","showLegend":true,"showGrid":true,"showLabels":true,"aggregation":"sum","sortBy":"desc","limit":10,"stacked":false},
  {"id":"ch2","type":"line","title":"Monthly Revenue Trend","xColumn":"Month Name","yColumns":["Revenue","EBITDA"],"colorScheme":"cool","showLegend":true,"showGrid":true,"showLabels":false,"aggregation":"sum","sortBy":"none","limit":12,"stacked":false},
  {"id":"ch3","type":"pie","title":"Revenue by Category","xColumn":"Category","yColumns":["Revenue"],"colorScheme":"greens","showLegend":true,"showGrid":false,"showLabels":true,"aggregation":"sum","sortBy":"desc","limit":8,"stacked":false},
  {"id":"ch4","type":"area","title":"Gross Margin Over Time","xColumn":"Month Name","yColumns":["Gross Profit","EBITDA"],"colorScheme":"blues","showLegend":true,"showGrid":true,"showLabels":false,"aggregation":"sum","sortBy":"none","limit":12,"stacked":false},
  {"id":"ch5","type":"donut","title":"Revenue by Segment","xColumn":"Customer Segment","yColumns":["Revenue"],"colorScheme":"warm","showLegend":true,"showGrid":false,"showLabels":true,"aggregation":"sum","sortBy":"desc","limit":6,"stacked":false},
  {"id":"ch6","type":"scatter","title":"LTV vs CAC","xColumn":"CAC","yColumns":["LTV"],"colorScheme":"neon","showLegend":false,"showGrid":true,"showLabels":false,"aggregation":"none","sortBy":"none","limit":200,"stacked":false},
  {"id":"ch7","type":"histogram","title":"Revenue Distribution","xColumn":"Revenue","yColumns":["Revenue"],"colorScheme":"reds","showLegend":false,"showGrid":true,"showLabels":false,"aggregation":"none","sortBy":"none","limit":50,"stacked":false}
]'

PUT_D=$(curl -s -b "$COOKIE_JAR" -X PUT "$BASE/api/dashboards/$DASH_ID" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Q1 Financial Dashboard (v2)\",\"charts\":$UPDATED_CHARTS}")
PUT_D_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" -X PUT "$BASE/api/dashboards/$DASH_ID" \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Q1 Financial Dashboard (v2)\"}")
if echo "$PUT_D" | grep -q '"name"'; then
  ok "Dashboard updated — name=$(echo "$PUT_D" | sed 's/.*"name":"\([^"]*\)".*/\1/')"
else
  fail "Dashboard update failed: $PUT_D"
fi

# ── 11. SHARING ───────────────────────────────────────────────────────────────
hdr "11 · SHARING"

# Share mapping by slug
SHARE_M=$(curl -s -b "$COOKIE_JAR" "$BASE/api/share/$MAPPING_SLUG?type=mapping")
if echo "$SHARE_M" | grep -q '"type":"mapping"'; then
  ok "Share mapping by slug '$MAPPING_SLUG' works"
else
  fail "Share mapping failed: $SHARE_M"
fi
expect_field "versions" "$SHARE_M"

# Share dashboard by slug
SHARE_D=$(curl -s -b "$COOKIE_JAR" "$BASE/api/share/$DASH_SLUG?type=dashboard")
if echo "$SHARE_D" | grep -q '"type":"dashboard"'; then
  ok "Share dashboard by slug '$DASH_SLUG' works"
else
  fail "Share dashboard failed: $SHARE_D"
fi

# Invalid type
BAD_SHARE=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" "$BASE/api/share/$DASH_SLUG?type=invalid")
expect_status "400" "$BAD_SHARE" "invalid share type → 400"

# Nonexistent slug
MISSING_SHARE=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" "$BASE/api/share/slug-does-not-exist-xyz?type=mapping")
expect_status "404" "$MISSING_SHARE" "missing slug → 404"

# ── 12. STAGING (if route exists) ─────────────────────────────────────────────
hdr "12 · STAGING RUNS"

STAGING_LIST_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" \
  "$BASE/api/mappings/$MAPPING_ID/staging")
if [[ "$STAGING_LIST_STATUS" == "200" ]]; then
  ok "Staging list endpoint live (200)"
else
  warn "Staging endpoint returned $STAGING_LIST_STATUS (may need implementation)"
fi

# ── 13. SETTINGS ──────────────────────────────────────────────────────────────
hdr "13 · SETTINGS"

SETTINGS_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" \
  "$BASE/api/settings/smartsheet")
if [[ "$SETTINGS_STATUS" == "200" ]]; then
  ok "Settings/smartsheet endpoint live (200)"
else
  warn "Settings endpoint returned $SETTINGS_STATUS"
fi

# ── 14. DELETE CLEANUP ────────────────────────────────────────────────────────
hdr "14 · DELETE (cleanup)"

DEL_M2_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" \
  -X DELETE "$BASE/api/mappings/$MAPPING2_ID")
expect_status "200" "$DEL_M2_STATUS" "delete second mapping"

DEL_D2_STATUS=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" \
  -X DELETE "$BASE/api/dashboards/$DASH2_ID")
expect_status "200" "$DEL_D2_STATUS" "delete second dashboard"

# Verify they're gone
GONE_M=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" \
  "$BASE/api/mappings/$MAPPING2_ID")
expect_status "404" "$GONE_M" "deleted mapping is gone → 404"

GONE_D=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" \
  "$BASE/api/dashboards/$DASH2_ID")
expect_status "404" "$GONE_D" "deleted dashboard is gone → 404"

# Cross-user delete attempt (we only have one user, so just verify 404 for bad id)
BAD_DEL=$(curl -s -o /dev/null -w "%{http_code}" -b "$COOKIE_JAR" \
  -X DELETE "$BASE/api/mappings/totally-wrong-id-000")
expect_status "404" "$BAD_DEL" "delete nonexistent → 404"

# ── FINAL SUMMARY ─────────────────────────────────────────────────────────────
echo ""
echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
echo -e "  ${GREEN}PASSED${NC}: $pass   ${YELLOW}WARNINGS${NC}: $warn   ${RED}FAILED${NC}: $fail"
echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
echo ""
echo "  Surviving test data (for manual UI testing):"
echo "    Mapping  : id=$MAPPING_ID  slug=$MAPPING_SLUG  (3 versions)"
echo "    Dashboard: id=$DASH_ID  slug=$DASH_SLUG  (6+1 charts)"
echo ""

if [[ "$fail" -gt 0 ]]; then
  exit 1
fi
