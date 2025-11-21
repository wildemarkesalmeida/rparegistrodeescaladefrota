# RPA Registro de Escala de Frota

Automação Playwright para registrar a escala diária de motoristas e veículos no sistema Sankhya.

## Pré-requisitos

- Node.js 18+ e npm instalados (`node -v` e `npm -v` para confirmar)
- Navegador Chrome instalado (usado pelo canal do Playwright)
- Planilha de escala diária no formato definido (ver `ESCALA DIARIA - 04-11.xlsx` como exemplo)

## Configuração do ambiente

1. Instale as dependências
   ```bash
   npm install
   npx playwright install
   ```
2. Copie `config/env.example` para `.env` na raiz e preencha:
   ```dotenv
   RPA_BASE_URL=http://transul.snk.ativy.com:50150/mge/system.jsp#app/...
   RPA_USERNAME=SEU_USUARIO
   RPA_PASSWORD=SUA_SENHA
   RPA_SCHEDULE_FILE=ESCALA DIARIA - 04-11.xlsx
   RPA_SCHEDULE_SHEET=14-11-2025
   RPA_KEEP_BROWSER_OPEN=true
   ```
3. Garanta que a planilha tenha as colunas `PLACA`, `MOTORISTA`, `TURNO`, `TIPO` e uma aba para cada data no formato `dd-mm-aaaa` (a automação usa o nome da aba para definir a data da escala).

## Executando o RPA

```bash
npm run rpa
```

O script:
- Faz login no Sankhya
- Seleciona o menu Escala Motoristas
- Cadastra a data do dia seguinte
- Itera por todas as linhas da planilha inserindo motoristas/veículos

Logs do terminal mostram o motorista/veículo atual. Se `RPA_KEEP_BROWSER_OPEN=true`, o navegador permanece aberto ao final (encerre com `Ctrl+C`).

## Gerando executável (Windows x64)

```bash
npm run build:exe
```

O comando cria `dist/rpa-escala.exe`, incorporando a versão atual do script. Ao executar o `.exe`, o `.env` e a planilha devem estar no mesmo diretório da aplicação.

## Solução de problemas

- **Timeout em campos de pesquisa**: confirme que o nome/placa na planilha coincide com o cadastro do Sankhya. O script tenta selecionar a sugestão; se não existir, confirma com Tab.
- **Mudança de layout**: atualize os seletores em `src/rpa.js` (funções `setTurno`, `setMotorista`, `setVeiculo`, `setTipo` e `createNewEntry`).
- **Planilha não encontrada**: ajuste `RPA_SCHEDULE_FILE` e `RPA_SCHEDULE_SHEET` no `.env`.

## Estrutura principal

- `src/rpa.js`: fluxo completo processando todos os motoristas da planilha
- `src/rpa_single.js`: versão de referência para cadastrar apenas o primeiro registro
- `config/env.example`: modelo de variáveis de ambiente
- `ESCALA DIARIA - 04-11.xlsx`: exemplo de planilha

## Licença

Projeto sob licença ISC (padrão `package.json`). Ajuste conforme políticas internas se necessário.

