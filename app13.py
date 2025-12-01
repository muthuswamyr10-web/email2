#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import imaplib, smtplib, email, webbrowser, os, threading, re, time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import decode_header
from datetime import datetime, timedelta

try: import openpyxl
except: openpyxl=None

try:
    from PIL import Image, ImageTk
except: Image=ImageTk=None

# -----------------------------
# CONFIG
# -----------------------------
YAHOO_EMAIL = "muthuswamy23@yahoo.in"
APP_PASSWORD = "oeuvcfzkvrlfjvul"
DEFAULT_REPLY = """Hi, 
Please find my CV/Resume attached.
Let me know if you need any additional information.

Regards,
Muthuswmy R
PH 8056256894 """

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
URL_RE = re.compile(r"https?://[^\s]+")

# -----------------------------
# UTILITIES
# -----------------------------
def decode_mime_words(value):
    if not value: return ""
    parts=[]
    for frag,enc in decode_header(value):
        if isinstance(frag,bytes):
            try: parts.append(frag.decode(enc or "utf-8", errors="ignore"))
            except: parts.append(frag.decode("utf-8", errors="ignore"))
        else: parts.append(frag)
    return "".join(parts)

def extract_emails_from_text(text): return EMAIL_RE.findall(text or "")
def extract_urls_from_text(text): return URL_RE.findall(text or "")
def clean_keyword_lines(text): return "\n".join([line.strip() for line in text.splitlines() if line.strip()])

def safe_get_body_text(msg):
    body=""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type()=="text/plain":
                    payload=part.get_payload(decode=True)
                    if payload:
                        charset=part.get_content_charset() or "utf-8"
                        try: body+=payload.decode(charset, errors="ignore")
                        except: body+=payload.decode("utf-8", errors="ignore")
        else:
            payload=msg.get_payload(decode=True)
            if payload:
                charset=msg.get_content_charset() or "utf-8"
                try: body=payload.decode(charset, errors="ignore")
                except: body=payload.decode("utf-8", errors="ignore")
    except: pass
    return body

def parse_internaldate(fetch_response):
    try:
        header_part=None
        if isinstance(fetch_response,tuple) and isinstance(fetch_response[0],bytes): header_part=fetch_response[0].decode(errors="ignore")
        elif isinstance(fetch_response,bytes): header_part=fetch_response.decode(errors="ignore")
        else: header_part=str(fetch_response)
        m=re.search(r'INTERNALDATE\s+"([^"]+)"',header_part)
        if m: return m.group(1)
    except: pass
    return datetime.utcnow().strftime("%d-%b-%Y %H:%M:%S +0000")

def imap_search_between(imap_conn,start_date,end_date):
    try:
        criteria=f'(SINCE "{start_date}" BEFORE "{end_date}")'
        typ,data=imap_conn.search(None,criteria)
        if typ!="OK" or not data: return []
        return data[0].split()
    except: return []

# -----------------------------
# FETCH EMAILS WORKER
# -----------------------------
def worker_fetch(start_date,end_date,keywords,progress_cb,stop_event,result_cb,error_cb):
    seen=set()
    try:
        imap=imaplib.IMAP4_SSL("imap.mail.yahoo.com")
        imap.login(YAHOO_EMAIL,APP_PASSWORD)
        imap.select("INBOX")
    except Exception as exc: error_cb(f"IMAP login failed: {exc}"); return

    try:
        ids=imap_search_between(imap,start_date,end_date)
        total=len(ids)
        for idx,e_id in enumerate(ids,start=1):
            if stop_event.is_set(): break
            progress_cb(idx,total)
            try:
                typ,msg_data=imap.fetch(e_id,"(RFC822 INTERNALDATE)")
                if typ!="OK" or not msg_data: continue

                internal_date=parse_internaldate(msg_data[0][0] if isinstance(msg_data[0],tuple) else msg_data[0])
                raw=msg_data[0][1] if isinstance(msg_data[0],tuple) else msg_data[0]
                if not raw: continue

                msg=email.message_from_bytes(raw)
                sender=decode_mime_words(msg.get("From",""))
                subject=decode_mime_words(msg.get("Subject",""))
                body=safe_get_body_text(msg)
                combined=(subject or "")+"\n"+(body or "")

                matched_kw=None
                for kw in keywords:
                    if kw.lower() in combined.lower(): matched_kw=kw; break

                if matched_kw:
                    emails_found=set(extract_emails_from_text(combined))
                    if not emails_found: emails_found.update(extract_emails_from_text(sender))
                    if not emails_found: emails_found.add("<no-email>")
                    urls=extract_urls_from_text(combined)
                    url=urls[0] if urls else ""
                    for fe in emails_found:
                        key=(sender.lower(),fe.lower(),internal_date)
                        if key in seen: continue
                        seen.add(key)
                        full_text=f"INTERNALDATE: {internal_date}\nFrom: {sender}\nSubject: {subject}\n\n{body}"
                        result_cb({"internal_date":internal_date,"sender":sender,"subject":subject,"found_email":fe,"keyword":matched_kw,"link":url,"full_text":full_text})
            except: continue
    finally:
        try: imap.logout()
        except: pass

# -----------------------------
# SEND EMAIL
# -----------------------------
def send_auto_reply(to_addr,original_subject,body,attachments=None):
    try:
        attachments=attachments or []
        msg=MIMEMultipart()
        msg["From"]=YAHOO_EMAIL
        msg["To"]=to_addr
        # Keep subject as original_subject (no added "Re:" unless desired)
        msg["Subject"]=(original_subject or "")
        msg.attach(MIMEText(body,"plain"))
        for file_path in attachments:
            if not file_path or not os.path.isfile(file_path): continue
            part=MIMEBase("application","octet-stream")
            with open(file_path,"rb") as f: part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",f'attachment; filename="{os.path.basename(file_path)}"')
            msg.attach(part)
        server=smtplib.SMTP_SSL("smtp.mail.yahoo.com",465)
        server.login(YAHOO_EMAIL,APP_PASSWORD)
        server.sendmail(YAHOO_EMAIL,to_addr,msg.as_string())
        server.quit()
        return True,None
    except Exception as exc: return False,str(exc)

# -----------------------------
# EXPORT EXCEL
# -----------------------------
def export_to_excel(treeview):
    if openpyxl is None:
        messagebox.showerror("Missing dependency","Install openpyxl: pip install openpyxl")
        return
    if not treeview.get_children():
        messagebox.showerror("Error","No data to export.")
        return
    file_path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel Files","*.xlsx")])
    if not file_path: return
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Results"
    ws.append(["Date","Sender","Email","Subject","Keyword","Link","Full Email"])
    for row in treeview.get_children():
        vals=treeview.item(row,"values") or ()
        full=treeview.item(row,"tags")[1] if len(treeview.item(row,"tags"))>1 else ""
        ws.append(list(vals)+[full])
    wb.save(file_path)
    messagebox.showinfo("Success",f"Exported to:\n{file_path}")

# -----------------------------
# MODERN APP CLASS
# -----------------------------
class App:
    def __init__(self,root):
        self.root=root
        root.title("CV/Resume Email Extractor")
        root.geometry("1300x750")
        root.configure(bg="#f0f0f0")
        self.default_attachment=None
        self.stop_event=threading.Event()

        # Header with gradient (preserved)
        self.header=tk.Canvas(root,height=60,highlightthickness=0)
        self.header.pack(fill="x")
        self.gradient_colors=["#4a90e2","#50e3c2"]
        self.animate_gradient(0)
        self.header_text=self.header.create_text(80,30,text="CV/Resume Email Extractor",fill="white",font=("Helvetica",18,"bold"),anchor="w")

        # Main frame splits left (original UI) and right (manual sender)
        main_frame = ttk.Frame(root)
        main_frame.pack(fill="both", expand=True, padx=8, pady=6)

        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side="left", fill="both", expand=True)

        # -----------------------------
        # RIGHT PANEL (REPLACED) - soft grey background (#f0f0f0)
        # -----------------------------
        right_frame = tk.Frame(main_frame, bg="#f0f0f0", bd=1, relief="flat")
        right_frame.pack(side="right", fill="y", padx=(8,0), pady=8)

        # Manual Sender header
        hdr = tk.Label(right_frame, text="Manual Email Sender", bg="#f0f0f0", font=("Helvetica",12,"bold"))
        hdr.pack(anchor="nw", padx=12, pady=(12,6))

        # Email ID
        lbl_email = tk.Label(right_frame, text="Email ID:", bg="#f0f0f0")
        lbl_email.pack(anchor="w", padx=12, pady=(6,2))
        self.manual_email_entry = ttk.Entry(right_frame, width=36)
        self.manual_email_entry.pack(padx=12, pady=(0,6))

        # Subject
        lbl_subject = tk.Label(right_frame, text="Subject:", bg="#f0f0f0")
        lbl_subject.pack(anchor="w", padx=12, pady=(6,2))
        self.manual_subject_entry = ttk.Entry(right_frame, width=36)
        self.manual_subject_entry.insert(0, "Applying for Job Opportunity")
        self.manual_subject_entry.pack(padx=12, pady=(0,6))

        # Message (medium size)
        lbl_message = tk.Label(right_frame, text="Message:", bg="#f0f0f0")
        lbl_message.pack(anchor="w", padx=12, pady=(6,2))
        self.manual_msg = tk.Text(right_frame, height=7, width=40, wrap="word", bd=1, relief="solid")
        self.manual_msg.insert("1.0", DEFAULT_REPLY)
        self.manual_msg.pack(padx=12, pady=(0,8))

        # Attachment info label
        self.right_attachment_label = tk.Label(right_frame, text="No default attachment", bg="#f0f0f0", fg="#333333")
        self.right_attachment_label.pack(anchor="w", padx=12, pady=(4,8))

        # Send Now button (keeps original simple styling)
        send_btn = ttk.Button(right_frame, text="Send Now", command=self.send_manual_email)
        send_btn.pack(padx=12, pady=(6,12), fill="x")

        # ------------------ ORIGINAL UI IN LEFT PANEL (unchanged) -------------------
        # Top frame: Dates, default attachments
        top=ttk.Frame(left_frame); top.pack(padx=12,pady=(8,8),fill="x")
        ttk.Label(top,text="Start Date (DD-MM-YYYY):").grid(row=0,column=0,sticky="w")
        self.start_entry=ttk.Entry(top,width=16); self.start_entry.grid(row=0,column=1,padx=6)
        ttk.Label(top,text="End Date (DD-MM-YYYY):").grid(row=0,column=2,sticky="w")
        self.end_entry=ttk.Entry(top,width=16); self.end_entry.grid(row=0,column=3,padx=6)
        today=datetime.today().strftime("%d-%m-%Y")
        self.start_entry.insert(0,today); self.end_entry.insert(0,today)

        ttk.Button(top,text="Set Default Attachment",command=self.select_default_attachment).grid(row=0,column=4,padx=5)
        ttk.Button(top,text="Remove Default Attachment",command=self.remove_default_attachment).grid(row=0,column=5,padx=5)
        self.default_attachment_label=ttk.Label(top,text="No default attachment",foreground="#333333")
        self.default_attachment_label.grid(row=0,column=6,padx=5)

        # Keywords
        ttk.Label(left_frame,text="Extra Keywords (one phrase per line):").pack(anchor="w", padx=12, pady=(8,0))
        self.kw_text=tk.Text(left_frame,height=5,wrap="word",font=("Helvetica",10),bg="#e8f0fe",fg="#333333",bd=2,relief="groove")
        self.kw_text.pack(fill="x", padx=12)

        # Buttons row
        btn_row=ttk.Frame(left_frame); btn_row.pack(padx=12,pady=8,fill="x")
        self.scan_btn=self.create_hover_button(btn_row,"Scan",self.start_scan)
        self.scan_btn.grid(row=0,column=0,padx=4)
        self.stop_btn=self.create_hover_button(btn_row,"Stop",self.stop_scan)
        self.stop_btn.grid(row=0,column=1,padx=4)
        self.export_btn=self.create_hover_button(btn_row,"Export to Excel",lambda: export_to_excel(self.tree))
        self.export_btn.grid(row=0,column=2,padx=4)

        # Progress bar
        self.progress=ttk.Progressbar(left_frame,length=700,style="TProgressbar"); self.progress.pack(padx=12,pady=6)

        # Treeview
        cols=("Date","Sender","Email","Subject","Keyword","Link","Reply")
        self.tree=ttk.Treeview(left_frame,columns=cols,show="headings",selectmode="browse")
        for c in cols:
            self.tree.heading(c,text=c)
            self.tree.column(c,width=140,anchor="w")
        self.tree.pack(expand=True,fill="both",padx=12,pady=8)
        self.tree.bind("<Double-1>",self.show_full_email)
        self.tree.tag_configure('oddrow',background='#f7f9fc')
        self.tree.tag_configure('evenrow',background='#ffffff')
        self.tree.bind("<Button-1>",self.tree_click)

    # ---------- HELPER METHODS ----------
    def animate_gradient(self,step):
        # Simple horizontal gradient animation
        self.header.delete("gradient")
        w=self.header.winfo_width() or 1300
        r1,g1,b1=self.hex_to_rgb(self.gradient_colors[0])
        r2,g2,b2=self.hex_to_rgb(self.gradient_colors[1])
        for i in range(w):
            r=int(r1 + (r2-r1)*(i/w))
            g=int(g1 + (g2-g1)*(i/w))
            b=int(b1 + (b2-b1)*(i/w))
            color=f"#{r:02x}{g:02x}{b:02x}"
            self.header.create_line(i,0,i,60,fill=color,tags="gradient")
        self.root.after(100,self.animate_gradient,(step+1)%100)

    def hex_to_rgb(self,h):
        h=h.lstrip("#")
        return tuple(int(h[i:i+2],16) for i in (0,2,4))

    def create_hover_button(self,parent,text,command):
        btn=ttk.Button(parent,text=text,command=command)
        def on_enter(e): btn.config(style="Hover.TButton")
        def on_leave(e): btn.config(style="TButton")
        btn.bind("<Enter>",on_enter)
        btn.bind("<Leave>",on_leave)
        return btn

    # Progress callback
    def progress_cb(self,current,total):
        self.progress["maximum"]=total if total>0 else 1
        self.progress["value"]=current

    # Result callback
    def result_cb(self,data):
        idx=len(self.tree.get_children())
        tag='evenrow' if idx%2==0 else 'oddrow'
        # store full email text as second tag so it can be retrieved later
        self.tree.insert("",tk.END,values=(data["internal_date"],data["sender"],data["found_email"],data["subject"],data["keyword"],data["link"],"Reply"),tags=(tag,data["full_text"]))

    def error_cb(self,msg): messagebox.showerror("Error",msg)

    # Treeview click
    def tree_click(self,event):
        region=self.tree.identify("region",event.x,event.y)
        if region=="cell":
            col=int(self.tree.identify_column(event.x).replace("#",""))-1
            row=self.tree.identify_row(event.y)
            if not row: return
            vals=self.tree.item(row,"values")
            tags=self.tree.item(row,"tags")
            if col==5:  # Link
                url=vals[5]
                if url: webbrowser.open(url)
            elif col==6: # Reply
                to_addr=vals[2] if vals[2] and vals[2]!="<no-email>" else vals[1]
                ok,err=send_auto_reply(to_addr,vals[3],DEFAULT_REPLY,[self.default_attachment] if self.default_attachment else [])
                if ok: messagebox.showinfo("Success",f"Email sent to {to_addr}")
                else: messagebox.showerror("Error",f"Failed:\n{err}")

    # -----------------------------
    # Double-click full email popup with editable reply
    # -----------------------------
    def show_full_email(self,event):
        sel=self.tree.focus()
        if not sel: return
        vals=self.tree.item(sel,"values")
        tags=self.tree.item(sel,"tags")
        full=tags[1] if len(tags)>1 else ""

        win=tk.Toplevel(self.root)
        win.title("Full Email")
        win.geometry("900x700")
        win.configure(bg="#f0f4f8")

        header=f"Date: {vals[0]}\nFrom: {vals[1]}\nEmail: {vals[2]}\nSubject: {vals[3]}\nKeyword: {vals[4]}\nLink: {vals[5]}\n\n"
        txt=tk.Text(win,wrap="word",font=("Helvetica",10),bg="#ffffff",fg="#333333",bd=2,relief="groove")
        txt.pack(expand=True,fill="both",padx=10,pady=10)
        txt.insert("1.0",header+full)

        def open_reply_editor():
            edit_win=tk.Toplevel(win)
            edit_win.title("Edit Auto Reply")
            edit_win.geometry("600x400")
            edit_win.configure(bg="#f0f4f8")

            ttk.Label(edit_win,text="Edit your auto-reply message:").pack(anchor="w",padx=6,pady=4)
            reply_box=tk.Text(edit_win,height=10,wrap="word",font=("Helvetica",10),bg="#e8f0fe",fg="#333333",bd=2,relief="groove")
            reply_box.pack(fill="both",padx=6,pady=4,expand=True)
            reply_box.insert("1.0",DEFAULT_REPLY)

            attachments=[]

            def add_attachment():
                file_path=filedialog.askopenfilename()
                if file_path:
                    attachments.append(file_path)
                    att_listbox.insert(tk.END,os.path.basename(file_path))

            def remove_attachment():
                sel_idx=att_listbox.curselection()
                if sel_idx:
                    idx=sel_idx[0]
                    attachments.pop(idx)
                    att_listbox.delete(idx)

            ttk.Button(edit_win,text="Add Attachment",command=add_attachment).pack(pady=4)
            ttk.Button(edit_win,text="Remove Selected Attachment",command=remove_attachment).pack(pady=4)
            att_listbox=tk.Listbox(edit_win,height=4,bg="#ffffff",fg="#333333",font=("Helvetica",9))
            att_listbox.pack(fill="x",padx=6,pady=2)

            ttk.Button(edit_win,text="Send Reply",command=lambda: self.send_reply_from_editor(vals,reply_box,attachments,edit_win)).pack(pady=6)

        ttk.Button(win,text="Reply...",command=open_reply_editor).pack(pady=6)

    def send_reply_from_editor(self,vals,reply_box,attachments,edit_win):
        body=reply_box.get("1.0","end").strip()
        to_addr=vals[2] if vals[2] and vals[2]!="<no-email>" else vals[1]
        ok,err=send_auto_reply(to_addr,vals[3],body,attachments)
        if ok:
            messagebox.showinfo("Success",f"Auto-reply sent to {to_addr}")
            edit_win.destroy()
        else:
            messagebox.showerror("Error",f"Failed to send:\n{err}")

    # -----------------------------
    # Manual send (right panel)
    # -----------------------------
    def send_manual_email(self):
        to_addr = self.manual_email_entry.get().strip()
        subject = self.manual_subject_entry.get().strip()
        body = self.manual_msg.get("1.0","end").strip()

        if not to_addr:
            messagebox.showerror("Error", "Please enter an email ID")
            return

        # basic email format check: if it looks invalid ask for confirmation
        if not EMAIL_RE.match(to_addr):
            if not EMAIL_RE.search(to_addr):
                if not messagebox.askyesno("Confirm", f"{to_addr} doesn't look like a valid email. Send anyway?"):
                    return

        attachments = [self.default_attachment] if self.default_attachment else []

        ok, err = send_auto_reply(to_addr, subject or "Apply to job", body, attachments)
        if ok:
            messagebox.showinfo("Success", f"Email sent to {to_addr}")
        else:
            messagebox.showerror("Error", f"Failed: {err}")

    # -----------------------------
    # Scan / Stop
    # -----------------------------
    def start_scan(self):
        try:
            s=datetime.strptime(self.start_entry.get().strip(),"%d-%m-%Y").strftime("%d-%b-%Y")
            e=datetime.strptime(self.end_entry.get().strip(),"%d-%m-%Y").strftime("%d-%b-%Y")
        except:
            messagebox.showerror("Error","Invalid date format (DD-MM-YYYY)"); return

        raw_kw=self.kw_text.get("1.0","end")
        cleaned=clean_keyword_lines(raw_kw)
        self.kw_text.delete("1.0","end"); self.kw_text.insert("1.0",cleaned)
        keywords=[k.strip() for k in cleaned.split("\n") if k.strip()]

        for i in self.tree.get_children(): self.tree.delete(i)
        self.stop_event.clear()
        threading.Thread(target=worker_fetch,args=(s,e,keywords,self.progress_cb,self.stop_event,self.result_cb,self.error_cb),daemon=True).start()

    def stop_scan(self):
        self.stop_event.set()
        try:
            self.progress["value"]=0
        except Exception:
            pass

    # -----------------------------
    # Default attachment
    # -----------------------------
    def select_default_attachment(self):
        file_path=filedialog.askopenfilename()
        if file_path:
            self.default_attachment=file_path
            self.default_attachment_label.config(text=os.path.basename(file_path))
            self.right_attachment_label.config(text=os.path.basename(file_path))
            messagebox.showinfo("Default Attachment",f"Set: {os.path.basename(file_path)}")

    def remove_default_attachment(self):
        self.default_attachment=None
        self.default_attachment_label.config(text="No default attachment")
        self.right_attachment_label.config(text="No default attachment")
        messagebox.showinfo("Default Attachment","Removed")

# -----------------------------
# MAIN
# -----------------------------
if __name__=="__main__":
    root=tk.Tk()

    style=ttk.Style()
    style.configure("TButton",padding=6)
    style.configure("Hover.TButton",padding=6,background="#4a90e2",foreground="white")

    App(root)
    root.mainloop()
