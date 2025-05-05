import smtplib
import csv
import re
import time
import random
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox, Text, END
from email.message import EmailMessage

# Email format validation
def is_valid_email(email):
    pattern = r"^[\w\.-]+@[\w\.-]+\.\w{2,}$"
    return re.match(pattern, email) is not None

# Load emails from CSV
def load_emails_from_csv(path):
    with open(path, newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        return list(set([row[0].strip() for row in reader if row and is_valid_email(row[0].strip())]))

# Send emails
def send_emails():
    email_address = email_entry.get()
    password = password_entry.get()
    subject = subject_entry.get()
    body = body_text.get("1.0", END)

    if not all([email_address, password, subject, body.strip(), cv_file_path.get(), motivation_file_path.get(), email_file_path.get()]):
        messagebox.showerror("Erreur", "Veuillez remplir tous les champs et sélectionner tous les fichiers.")
        return

    try:
        email_list = load_emails_from_csv(email_file_path.get())
        count_sent = 0
        total = len(email_list)

        if total == 0:
            messagebox.showwarning("Aucun email", "Aucun email valide trouvé dans le fichier CSV.")
            return

        for email in email_list:
            if count_sent >= 490:
                messagebox.showinfo("Limite atteinte", "490 emails envoyés. Arrêt automatique.")
                break

            msg = EmailMessage()
            msg['Subject'] = subject
            msg['From'] = email_address
            msg['To'] = email
            msg.set_content(body)

            # Attachments
            with open(cv_file_path.get(), 'rb') as f:
                msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename='CV.pdf')
            with open(motivation_file_path.get(), 'rb') as f:
                msg.add_attachment(
                    f.read(),
                    maintype='application',
                    subtype='vnd.openxmlformats-officedocument.wordprocessingml.document',
                    filename='Motivation_Letter.docx'
                )

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(email_address, password)
                smtp.send_message(msg)

            count_sent += 1
            progress_label.config(text=f"Progression : {count_sent} / {total} emails envoyés")
            root.update()

            # Pause aléatoire entre chaque email
            time.sleep(random.uniform(5, 15))

            # Pause prolongée toutes les 50 candidatures
            if count_sent % 50 == 0:
                pause_duration = random.uniform(60, 120)
                progress_label.config(text=f"Pause prolongée ({int(pause_duration)}s)... ({count_sent} envoyés)")
                root.update()
                messagebox.showinfo(
                    "Pause automatique",
                    f"Une pause de {int(pause_duration)} secondes est effectuée pour éviter les filtres anti-spam.\n\n{count_sent} emails ont été envoyés jusqu'à présent."
                )
                time.sleep(pause_duration)

        messagebox.showinfo("Terminé", f"{count_sent} emails envoyés avec succès !")

    except Exception as e:
        messagebox.showerror("Erreur", str(e))

# UI setup
root = Tk()
root.title("Email Sender - Stage Candidature")
root.geometry("600x700")

Label(root, text="Votre email Gmail :").pack()
email_entry = Entry(root, width=60)
email_entry.pack()

Label(root, text="Mot de passe d'application :").pack()
password_entry = Entry(root, show='*', width=60)
password_entry.pack()

Label(root, text="Sujet de l'email :").pack()
subject_entry = Entry(root, width=60)
subject_entry.pack()

Label(root, text="Corps de l'email :").pack()
body_text = Text(root, height=10, width=60)
body_text.pack()

cv_file_path = Entry(root, width=60)
motivation_file_path = Entry(root, width=60)
email_file_path = Entry(root, width=60)

Button(root, text="Sélectionner CV (PDF)", command=lambda: cv_file_path.insert(0, filedialog.askopenfilename())).pack()
cv_file_path.pack()

Button(root, text="Sélectionner Lettre de motivation (DOCX)", command=lambda: motivation_file_path.insert(0, filedialog.askopenfilename())).pack()
motivation_file_path.pack()

Button(root, text="Sélectionner fichier d'emails (CSV)", command=lambda: email_file_path.insert(0, filedialog.askopenfilename())).pack()
email_file_path.pack()

progress_label = Label(root, text="Progression : 0 / 0 emails envoyés", fg="blue")
progress_label.pack(pady=10)

Button(root, text="Envoyer les emails", command=send_emails, bg="green", fg="white", padx=10, pady=5).pack(pady=20)

root.mainloop()
