import pandas as pd
import re
import os
from jpype import startJVM, shutdownJVM, JClass, getDefaultJVMPath, java, isJVMStarted

def setup_java():
    if isJVMStarted():
        return JClass("com.cybozu.labs.langdetect.DetectorFactory")

    java_home = os.environ.get('JAVA_HOME')
    if not java_home:
        raise EnvironmentError("JAVA_HOME is not set")
    
    jvm_path = os.path.join(java_home, 'lib', 'server', 'libjvm.dylib')
    if not os.path.exists(jvm_path):
        jvm_path = os.path.join(java_home, 'jre', 'lib', 'server', 'libjvm.dylib')
    if not os.path.exists(jvm_path):
        raise FileNotFoundError(f"Cannot find libjvm.dylib in {java_home}")
    
    # Kullanıcıdan language-detection dizininin yolunu alın
    lang_detect_dir = input("Lütfen language-detection dizininin tam yolunu girin: ").strip()
    jar_path = os.path.join(lang_detect_dir, "dist", "language-detection.jar")
    profile_dir = os.path.join(lang_detect_dir, "profiles")
    
    if not os.path.exists(jar_path):
        raise FileNotFoundError(f"Cannot find language-detection.jar in {jar_path}")
    
    startJVM(jvm_path, "-Djava.class.path=" + jar_path)
    
    DetectorFactory = JClass("com.cybozu.labs.langdetect.DetectorFactory")
    DetectorFactory.loadProfile(profile_dir)
    
    return DetectorFactory

def is_turkish_email(email, detector_factory):
    turkish_domains = ['.tr']
    
    if any(domain in email.lower() for domain in turkish_domains):
        return True
    
    parts = email.split('@')
    if len(parts) != 2:
        return False
    
    username, domain = parts
    domain = domain.split('.')[0]
    
    try:
        detector = detector_factory.create()
        detector.append(username + " " + domain)
        language = detector.detect()
        return language == 'tr'
    except java.lang.Exception as e:
        print(f"Dil tespiti hatası ({email}): {e}")
        return False

def extract_and_categorize_emails(input_file, output_file):
    detector_factory = None
    try:
        detector_factory = setup_java()

        df = pd.read_excel(input_file)
        all_data = df.values.ravel()

        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        turkish_emails = set()
        foreign_emails = set()

        for item in all_data:
            if isinstance(item, str):
                found_emails = re.findall(email_pattern, item)
                for email in found_emails:
                    if is_turkish_email(email, detector_factory):
                        turkish_emails.add(email)
                    else:
                        foreign_emails.add(email)

        turkish_df = pd.DataFrame(list(turkish_emails), columns=["Türk E-posta Adresleri"])
        foreign_df = pd.DataFrame(list(foreign_emails), columns=["Yabancı E-posta Adresleri"])

        with pd.ExcelWriter(output_file) as writer:
            turkish_df.to_excel(writer, sheet_name="Türk E-postalar", index=False)
            foreign_df.to_excel(writer, sheet_name="Yabancı E-postalar", index=False)

        print(f"{len(turkish_emails)} adet Türk e-posta adresi ve {len(foreign_emails)} adet yabancı e-posta adresi bulundu.")
        print(f"Sonuçlar {output_file} dosyasına kaydedildi.")

    except Exception as e:
        print(f"Bir hata oluştu: {e}")
    finally:
        if isJVMStarted():
            shutdownJVM()

if __name__ == "__main__":
    input_file = input("Lütfen giriş Excel dosyasının adını girin (örn: Fuar Mail Listesi.xlsx): ").strip()
    output_file = input("Lütfen çıkış Excel dosyasının adını girin (örn: Ayrılmış E-posta Adresleri.xlsx): ").strip()
    extract_and_categorize_emails(input_file, output_file)
