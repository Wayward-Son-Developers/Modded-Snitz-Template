MOD: Custom Policy v1.2
[CREATE]
CUSTOM_POLICY
CP_ID
CP_MODE#int#NOT NULL#
CP_CONTENT#memo#NULL#
[END]
[INSERT]
CUSTOM_POLICY
(CP_MODE,CP_CONTENT)#(0,'Your custom policy here')
[END]
