import subprocess
import os

def login_existe_ad(login):
    comando = f'''
    $user = Get-ADUser -Filter {{ SamAccountName -eq "{login}" }} -ErrorAction SilentlyContinue
    if ($user) {{ Write-Output "existe" }}
    '''
    resultado = subprocess.run(["powershell", "-Command", comando], capture_output=True, text=True)
    return "existe" in resultado.stdout.lower()

def email_existe_ad(email):
    comando = f'''
    $user = Get-ADUser -Filter {{ EmailAddress -eq "{email}" }} -ErrorAction SilentlyContinue
    if ($user) {{ Write-Output "existe" }}
    '''
    resultado = subprocess.run(["powershell", "-Command", comando], capture_output=True, text=True)
    return "existe" in resultado.stdout.lower()

def gerar_login_email(nome_completo):
    particulas = {"de", "da", "do", "das", "dos", "e", "d’", "d."}

    partes = nome_completo.strip().lower().split()
    partes_sem_particulas = [p for p in partes if p not in particulas]

    primeiro_nome = partes_sem_particulas[0]
    ultimo_nome_real = partes_sem_particulas[-1]

    base_login = f"{primeiro_nome}_{ultimo_nome_real[0]}"
    login = base_login
    i = 1

    while login_existe_ad(login):
        if i < len(ultimo_nome_real):
            login = f"{primeiro_nome}_{ultimo_nome_real[:i+1]}"
            i += 1
        else:
            if len(partes_sem_particulas) > 2:
                meio = partes_sem_particulas[1]
                login = f"{primeiro_nome}_{meio[0]}{ultimo_nome_real[0]}"
            else:
                login = f"{primeiro_nome}_{ultimo_nome_real}"
            if not login_existe_ad(login):
                break

    dominio = "@tseaenergia.com.br"
    email_base = f"{primeiro_nome}.{ultimo_nome_real}"
    email = f"{email_base}{dominio}"

    if email_existe_ad(email):
        if len(partes_sem_particulas) > 2:
            segundo_sobrenome = partes_sem_particulas[-2]
            email_base = f"{primeiro_nome}.{segundo_sobrenome}"
            email = f"{email_base}{dominio}"

            if email_existe_ad(email):
                contador = 1
                while True:
                    tentativa = f"{email_base}{contador}{dominio}"
                    if not email_existe_ad(tentativa):
                        email = tentativa
                        break
                    contador += 1
        else:
            contador = 1
            while True:
                tentativa = f"{email_base}{contador}{dominio}"
                if not email_existe_ad(tentativa):
                    email = tentativa
                    break
                contador += 1

    return login, email

def criar_usuario_ad(nome_completo, usuario_modelo, senha_padrao="Mudar@2025", simular=True):
    login_novo, email = gerar_login_email(nome_completo)
    nome, sobrenome = nome_completo.strip().split(" ", 1)

    comando = f'''
    $modelo = Get-ADUser -Identity "{usuario_modelo}"

    New-ADUser -Name "{nome_completo}" -GivenName "{nome}" -Surname "{sobrenome}" `
        -SamAccountName "{login_novo}" `
        -UserPrincipalName "{email}" `
        -EmailAddress "{email}" `
        -DisplayName "{nome_completo}" `
        -Path $modelo.DistinguishedName.Substring($modelo.DistinguishedName.IndexOf(",")+1) `
        -AccountPassword (ConvertTo-SecureString "{senha_padrao}" -AsPlainText -Force) `
        -Enabled $true
        -ChangePasswordAtLogon $true

    $grupos = Get-ADUser "{usuario_modelo}" -Properties MemberOf | Select-Object -ExpandProperty MemberOf
    foreach ($grupo in $grupos) {{
        Add-ADGroupMember -Identity $grupo -Members "{login_novo}"
    }}
    '''
    print(f"\n--- COMANDO POWERHELL ---\n{comando}\n")

    if simular:
        print("🔍 MODO SIMULAÇÃO ATIVO: Nenhum usuário foi criado.")
        return login_novo, email

    confirmar = input("❗ Confirmar criação do usuário? [S/N]: ").strip().lower()
    if confirmar != "s":
        print("🚫 Criação cancelada pelo usuário.")
        return login_novo, email

    resultado = subprocess.run(["powershell", "-Command", comando], capture_output=True, text=True)

    if resultado.returncode == 0:
        print(f"\n✅ Usuário criado com sucesso!")
    else:
        print("\n❌ Erro ao criar usuário:\n", resultado.stderr)

    return login_novo, email

def formatar_nome(nome):
    palavras_minusculas = {"da", "de", "do", "das", "dos", "e"}
    palavras = nome.lower().split()
    nome_formatado = []
    for palavra in palavras:
        if palavra in palavras_minusculas:
            nome_formatado.append(palavra)
        else:
            nome_formatado.append(palavra.capitalize())
    return " ".join(nome_formatado)

def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

# --- Execução ---
if __name__ == "__main__":

    while True:
        clear()
        print("==== CRIAÇÃO DE USUÁRIO AD - TSEA ====")
        nome_completo = input("\nDigite o nome completo do novo colaborador: ").strip()
        nome_completo = formatar_nome(nome_completo)

        usuario_modelo = input("Digite o login do usuário modelo (ex: joao_p): ").strip()
        modo = input("Rodar em modo simulação? [S/N]: ").strip().lower()
        simular = True if modo == "s" else False

        login_gerado, email_gerado = criar_usuario_ad(nome_completo, usuario_modelo, simular=simular)

        print(f"\nResumo:")
        print(f"Usuário: {nome_completo}")
        print(f"Login: {login_gerado}")
        print(f"E-mail: {email_gerado}")
        print(f"Senha: Mudar@2025")

        continuar = input("\nDeseja cadastrar outro usuário? [S/N]: ").strip().lower()
        if continuar != "s":
            print("✅ Encerrando o programa. Até mais!")
            break
