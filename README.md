# RJP Requisições IP V4.0

Projeto único para **WebApp/PWA**, **Android APK** e **iOS**, baseado na aplicação HTML V3.0 enviada.

## GitHub
1. Criar repositório `RJP_Requisicoes_IP` no GitHub.
2. Carregar todos os ficheiros deste ZIP para a raiz do repositório (incluindo `.github/workflows`).
3. Em **Settings → Pages → Build and deployment**, escolher **GitHub Actions**.
4. O push para `main` executa **Publicar WebApp** e **Android APK**. A APK fica em **Actions → Android APK → Artifacts**.
5. O workflow **iOS Build** pode ser iniciado manualmente; valida o build para simulador, mas não produz um IPA instalável/assinado.

## Desenvolvimento local
`npm install` e `npm run dev` para WebApp. Para Android: `npm run build`, `npx cap add android`, `npx cap sync android`. Para iOS (macOS + Xcode): `npm run build`, `npx cap add ios`, `npx cap sync ios`, `npx cap open ios`.

## Dados e exportação
- Os registos continuam guardados localmente pelo código original V3.0 (`localStorage`). Use **Dados e segurança → Exportar cópia JSON** regularmente.
- Botão **Excel original** usa a estrutura do ficheiro `Requisições.xlsm` fornecido e exporta `.xlsx` sem macros; os campos são preenchidos e os totais usam fórmulas locais. Algumas células de identificação podem necessitar de ajuste à disposição exata do formulário original.
- Botão **PDF** abre a vista de impressão, onde se escolhe **Guardar como PDF**; no Android/iOS a disponibilidade depende do diálogo de impressão do sistema.
- A biblioteca Excel é obtida por CDN na primeira utilização; **exportação Excel não está garantida offline** sem incluir uma cópia local da biblioteca. O restante núcleo da app funciona offline depois de instalada.
- `templates/catalogo-referencia.xlsx` é mantido para referência; a pesquisa usa o catálogo embutido no HTML V3.0.

## Limitações a validar
Este é um projeto inicial pronto para carregar no GitHub, **não uma APK/IPA já compilada**. O Excel replica a folha base do modelo original, não garante equivalência total de macros, ligações externas e regras avançadas. Testar o documento com requisições reais antes de utilização operacional.

## V4.1 — Pesquisa por palavras

Na secção Requisições, pesquisar por várias palavras em qualquer ordem, sem acentos ou distinção de maiúsculas. A pesquisa abrange dados gerais e códigos/designações dos artigos de cada requisição. Inclui contador de resultados e botão para limpar. No editor, a pesquisa do catálogo LPU também aceita várias palavras. Tudo funciona localmente/offline.


## iOS V4.2 — IPA sem assinatura

O workflow `.github/workflows/ios.yml` compila para **iPhone físico** (`iphoneos`), empacota `Payload/App.app` e disponibiliza um `.ipa` **sem assinatura** nos Artifacts da execução GitHub Actions. Não é um IPA pronto a instalar diretamente: exige assinatura por uma ferramenta compatível (por exemplo, SideStore/AltStore/Sideloadly) e conta Apple, sujeito às limitações de assinatura gratuita. Não há garantia de instalação em todos os dispositivos.

GitHub > Actions > **iOS IPA (sem assinatura)** > Run workflow > Artifacts. Descarregar o ZIP do artifact, extrair o IPA e assinar/instalar com a ferramenta escolhida.


## Correção iOS V4.3
O workflow instala CocoaPods com `pod install --repo-update` em `ios/App` e compila `ios/App/App.xcworkspace` em vez de `App.xcodeproj`, para resolver o erro `unable to resolve module dependency: Capacitor`. O IPA é **não assinado** e necessita de assinatura para instalação. A compilação só pode ser validada no runner macOS do GitHub Actions.


## V4.4 — Ícone IP no telemóvel
Ícone gerado a partir do logótipo oficial da IP mostrado pelo utilizador na captura de ecrã. Inclui imagens para iOS, Android e PWA. Os workflows aplicam automaticamente os ícones após a criação do projeto Capacitor. O IPA continua sem assinatura e necessita de assinatura para instalação.


## V4.5 — correção ortográfica
Corrigida no catálogo de tipos de ação a expressão «Tipo de intervcenção / ação:» para «Tipo de intervenção / ação:» em `www/index.html`.


## V4.7
- Centros de Custo configuráveis: Caldas da Rainha, Coimbra, Guarda e Entroncamento; permite adicionar Aveiro, Porto e outros.
- Créditos: Rui Jorge Pedro e Abílio Felix.
- PDF A4 multipágina, com acentuação portuguesa, artigos, valores e total, para guardar/partilhar/imprimir.
- Numeração sequencial continua **local por instalação**. Não utilizar em simultâneo em dispositivos diferentes como numeração contratual definitiva sem backend central.
- A impressão iOS/Android depende do visualizador PDF do dispositivo.
- Os dados ficam no armazenamento local: exporte cópias de segurança.


## V4.8 — Rascunhos
- Guardar rascunhos incompletos, mesmo sem Centro de Custo, data ou artigos.
- Reabrir e continuar a editar pelo histórico.
- Número da requisição editável manualmente; atribuição automática ao finalizar se ficar vazio.
- Mensagem explícita se o armazenamento local falhar.
- A numeração e os dados permanecem locais a cada dispositivo; faça cópias de segurança antes de reinstalar.
