# Guia rápido de estrutura para o app financeiro em Android

O snapshot abaixo corresponde à estrutura padrão do Android Studio (módulo único `app`) semelhante à que você exibiu no print. Use este guia para saber exatamente onde criar cada pacote e arquivo sugerido no plano anterior.

## Onde ficam os arquivos Kotlin
1. No painel **Project** do Android Studio, altere a visualização para **Android** (como no print). O caminho base é `app/java/com.example.cadern…` (o namespace exibido no print começa com `com.example.cadern`).
2. Dentro desse pacote, crie as pastas a seguir (botão direito ➜ *New* ➜ *Package*). Ficam todas em `app/src/main/java/com.example.cadern…/`:
   - `data/google` → `SheetsServiceFactory.kt` (fábrica do cliente Sheets/Drive carregando a chave JSON via `assets`).
   - `data/sheets` → `FinanceSheetRepository.kt` (operações na planilha: inserir/atualizar linhas, ler mês/ano).
   - `domain/usecase` → `RecordDailySales.kt`, `RecordInvestment.kt`, `GetMonthlySummary.kt`, `GetYearlySummary.kt` (regras de negócio e agregações).
   - `ui/home` → `HomeViewModel.kt`, `HomeScreen.kt` (card de lucro do mês + botões).
   - `ui/dialog` → `RecordSalesDialog.kt`, `RecordInvestmentDialog.kt` (pop-ups com campos opcionais para vendas/investimento).
   - `ui/sheet` → `SheetViewModel.kt`, `SheetScreen.kt` (lista da planilha completa).
   - `ui/analytics` → `AnalyticsViewModel.kt`, `AnalyticsScreen.kt` (gráficos + seletor de mês/ano).
   - `ui/theme` → `Theme.kt`, `Color.kt`, `Type.kt` (tema escuro, verde para lucro positivo e vermelho para negativo).

> Dica: se preferir centralizar navegação, crie também `ui/navigation` para o `NavHost` que liga Home, Planilha e Análises.

## Recursos (XML) em `app/src/main/res/`
- **Tema/cores:** `res/values/themes.xml` e `res/values/colors.xml` já existem; adicione as cores escuras e tons de verde/vermelho. Se usar Compose puro, ainda vale definir cores no `ui/theme`, mas mantenha os arquivos XML sincronizados.
- **Strings:** `res/values/strings.xml` para textos de botões (“Adicionar vendas”, “Adicionar investimento”, etc.).
- **Ícones/imagens:** `res/drawable` ou `res/mipmap-*` (os ícones do launcher já estão em `mipmap`).

## Dependências no Gradle
Edite `app/build.gradle.kts` (módulo `app`) e inclua:
- Compose/Material3 (se ainda não estiver): `androidx.compose.*`, `androidx.activity:activity-compose`.
- APIs do Google: `com.google.apis:google-api-services-sheets` e `com.google.auth:google-auth-library-oauth2-http`.
- Gráficos: `com.github.PhilJay:MPAndroidChart` ou alternativa Compose.
Garanta repositórios `google()` e `mavenCentral()` no `settings.gradle.kts`/`build.gradle.kts` do projeto.

## Chave da service account
Para desenvolvimento local, coloque o JSON em `app/src/main/assets/finance_service_account.json` (crie a pasta `assets` se não existir). No `SheetsServiceFactory`, carregue com `context.assets.open("finance_service_account.json")`. Evite commitar a chave em repositório público.

## Sequência sugerida de implementação
1. Criar os pacotes acima em `java/com.example.cadern…`.
2. Ajustar dependências no Gradle do módulo `app`.
3. Implementar `SheetsServiceFactory` e `FinanceSheetRepository` (camada de dados na nuvem, sem cache local).
4. Implementar casos de uso em `domain/usecase`.
5. Construir a `HomeScreen` com o card de lucro + botões que abrem os diálogos.
6. Criar diálogos de vendas/investimento (campos opcionais, botões Voltar/Confirmar).
7. Implementar telas de Planilha e Análises (com seletor de mês/ano e gráficos mensais/anuais).
8. Finalizar o tema escuro em `ui/theme` e garantir cores verde/vermelho para o lucro.

Assim você consegue localizar rapidamente onde cada parte do código deve ser criada dentro da árvore padrão exibida no Android Studio.
