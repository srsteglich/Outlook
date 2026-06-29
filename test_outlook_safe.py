# test_outlook_safe.py
import win32com.client
import pythoncom
import re

def extract_email_from_ex(ex_address):
    """
    Extrai o email de um endereço EX do Exchange
    """
    if not ex_address:
        return None
    
    # Tenta encontrar qualquer email no texto
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    match = re.search(email_pattern, ex_address)
    if match:
        return match.group(0)
    
    # Tenta extrair do formato GUID do Exchange
    # Exemplo: /o=ExchangeLabs/ou=Exchange Administrative Group (FYDI0HF23SPDLT)/cn=Recipients/cn=be18e42009ca47e4838654d1b7c305b8-e955ca44-38
    # Não tem email neste formato, precisa buscar via Exchange
    return None

def get_outlook_email_safe():
    """
    Converte endereço EX para SMTP
    """
    email = None
    
    try:
        print("Conectando ao Outlook...")
        
        # Inicializa COM
        pythoncom.CoInitialize()
        
        # Conecta ao Outlook
        outlook = None
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            _ = outlook.Name
            print("✅ Outlook conectado com sucesso!")
        except Exception as e:
            print(f"❌ Erro ao conectar ao Outlook: {e}")
            return None
        
        # Obtém namespace
        try:
            # O namespace MAPI é a interface principal usada para acessar as configurações, pastas e itens da sua caixa de correio
            namespace = outlook.GetNamespace("MAPI")
            print("✅ Namespace MAPI obtido")
        except Exception as e:
            print(f"❌ Erro ao obter namespace: {e}")
            return None
        
        # Tenta obter contas primeiro
        try:
            print("\n📧 Contas configuradas no Outlook:")
            for account in namespace.Accounts:
                print(f"  - {account.DisplayName}")
                try:
                    if account.CurrentUser:
                        # Tenta obter o endereço
                        address_entry = account.CurrentUser.AddressEntry
                        if address_entry:
                            print(f"    Tipo: {address_entry.Type}")
                            
                            # Se for EX, converte para SMTP
                            if address_entry.Type == "EX":
                                try:
                                    exchange_user = address_entry.GetExchangeUser()
                                    if exchange_user:
                                        email = exchange_user.PrimarySmtpAddress
                                        print(f"    ✅ Email (Exchange): {email}")
                                except Exception as e:
                                    print(f"    Erro ao converter Exchange: {e}")
                            elif address_entry.Type == "SMTP":
                                email = address_entry.Address
                                print(f"    ✅ Email (SMTP): {email}")
                            else:
                                # Tenta outros métodos
                                email_candidate = account.CurrentUser.Address
                                if email_candidate and '@' in email_candidate:
                                    email = email_candidate
                                    print(f"    ✅ Email: {email}")
                            
                            if email:
                                break
                except Exception as e:
                    print(f"    Erro ao obter email desta conta: {e}")
        except Exception as e:
            print(f"Erro ao listar contas: {e}")
        
        # Se não encontrou email, tenta via CurrentUser
        if not email:
            try:
                current_user = namespace.CurrentUser
                print(f"\n👤 Usuário atual: {current_user}")
                
                # Tenta via AddressEntry
                if hasattr(current_user, 'AddressEntry'):
                    try:
                        address_entry = current_user.AddressEntry
                        if address_entry:
                            print(f"AddressEntry Type: {address_entry.Type}")
                            if address_entry.Type == "EX":
                                exchange_user = address_entry.GetExchangeUser()
                                if exchange_user:
                                    email = exchange_user.PrimarySmtpAddress
                                    print(f"✅ Email obtido via Exchange: {email}")
                            elif address_entry.Type == "SMTP":
                                email = address_entry.Address
                                print(f"✅ Email obtido via SMTP: {email}")
                    except Exception as e:
                        print(f"Erro ao processar AddressEntry: {e}")
                
                # Se ainda não tem email, tenta propriedades diretas
                if not email:
                    if hasattr(current_user, 'SMTPAddress'):
                        email = current_user.SMTPAddress
                        if email:
                            print(f"✅ Email via SMTPAddress: {email}")
                    
                    if not email and hasattr(current_user, 'Address'):
                        raw = current_user.Address
                        if raw and '@' in raw:
                            email = raw
                            print(f"✅ Email via Address: {email}")
                        
            except Exception as e:
                print(f"Erro ao processar CurrentUser: {e}")
        
        # Limpa recursos
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        
        # Valida se é um email válido
        if email and '@' in email:
            return email.lower()
        else:
            print(f"\n⚠️ Email inválido ou não encontrado: {email}")
            return None
        
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return None

if __name__ == "__main__":
    print("=" * 50)
    print("🔍 TESTE DE CONEXÃO COM OUTLOOK")
    print("=" * 50)
    print("Certifique-se que o Outlook está ABERTO")
    print("=" * 50)
    
    email = get_outlook_email_safe()
    
    print("\n" + "=" * 50)
    if email:
        print(f"✅ E-mail encontrado: {email}")
        print("✅ Formato correto para autenticação!")
    else:
        print("❌ Não foi possível obter o e-mail")
        print("\nPossíveis soluções:")
        print("1. Certifique-se que o Outlook está aberto")
        print("2. Execute como administrador")
        print("3. Verifique se há uma conta configurada no Outlook")
        print("4. Se estiver em rede corporativa, verifique as permissões")
    print("=" * 50)
    
    # Pausa para ver o resultado
    input("\nPressione ENTER para sair...")
