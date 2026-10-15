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
