# วิธีเตรียมเครื่อง Windows สำหรับการสอบ

1. แตก ZIP ไปยังโฟลเดอร์ standalone ใหม่ ห้ามวางใน production project หรือโฟลเดอร์ที่ sync กับ Drive
2. เปิด ChatGPT desktop app > Codex > Add/Open project (`Ctrl+O`) แล้วเลือกเฉพาะ `candidate_workspace`
3. ใช้ Windows native/PowerShell และตั้งสิทธิ์ Ask for approval หรือ workspace-only; ไม่ใช้ Full Access หากไม่จำเป็น
4. ห้ามเปิด production workspace, Drive connector, Gmail, Google Sheets, browser login หรือ credentials
5. ตรวจว่าไม่มี symlink, junction, reparse point หรือ absolute path ออกนอก workspace
6. อ่าน `INTERVIEW_CASE_TH.docx` และ `README_SAFETY.md` แล้วเริ่มด้วย `START_PROMPT_MINIMAL.txt`
7. เก็บ confidential interviewer package ไว้นอกเครื่องสอบและนอก Candidate ZIP
