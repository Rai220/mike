# Проверка подключения Microsoft 365

Это локальная диагностика Entra credentials и delegated доступа к почте. Она не подключает корпоративные данные к чату Mike, не запускает LLM и не изменяет RnD deployment.

## Проверенный результат на 2026-09-07

Client ID, Tenant ID и client secret сохранены в `backend/.env.microsoft365.local`. Файл исключён из Git, права `600`. Секрет и токены не включаются в отчёт.

Реальный `POST /{tenant}/oauth2/v2.0/token` с `grant_type=client_credentials` и Graph `.default` вернул **HTTP 200** и Bearer token. Microsoft принимает учётные данные приложения. Это не подтверждает доступ к пользовательской почте: в диагностически прочитанном токене приложения отсутствовали `scp` и application roles. Отсутствие ролей в app token не доказывает отсутствие настроенных delegated permissions.

Первая попытка входа завершилась `AADSTS50011` (`2026-09-07T06:26:25Z`) из-за несовпадения callback URI. После исправления и повторного входа проверка успешно завершилась:

```json
{"ok":true,"granted_scopes":["User.Read","Mail.Read"],"token_http_status":200,"profile_http_status":200,"mail_http_status":200,"message_count":1,"body_read":true,"mailbox_empty":false}
```

Подтверждены delegated авторизация, профиль и чтение тела одного письма. Содержимое письма и токены не сохранялись; диагностический процесс завершился. Точный тариф, административные настройки Conditional Access, refresh flow и постоянное подключение через Mike этим тестом не проверены. Первое приложение подключения описано в [инструкции настройки](microsoft365-setup.md).

## Настройка локальной проверки почты

В Microsoft Entra → App registrations → нужное приложение:

1. Authentication → Add a platform → **Web**: зарегистрировать `http://localhost:8400/microsoft365/callback` как redirect URI для локального теста. Не выбирать SPA; не добавлять тот же localhost path под другой тип приложения на другом порту.
2. API permissions → Microsoft Graph → **Delegated permissions**: `User.Read` и `Mail.Read`. Consent проходит при входе, либо его выдаёт администратор, если этого требует политика компании. Application `Mail.Read` и права отправки/записи для этого теста не нужны.
3. Входить рабочим аккаунтом соответствующего tenant, имеющим Exchange Online mailbox. Проверить назначение пользователя приложению, если оно требуется.

Это диагностический localhost callback, а не будущий production URL. Для корпоративного развёртывания потребуется отдельный HTTPS callback Mike, привязанный к авторизованной сессии приложения.

Локальный секретный файл содержит `MICROSOFT365_CLIENT_ID`, `MICROSOFT365_TENANT_ID`, `MICROSOFT365_CLIENT_SECRET` и справочный `MICROSOFT365_CLIENT_SECRET_ID`. Secret ID не участвует в OAuth token exchange. Скрипт читает файл явно; обычный backend пока не использует эти настройки.

## Запуск из корня репозитория

Требуется Node.js 22 или новее. Дополнительные зависимости не нужны.

```bash
node backend/scripts/microsoft365-check.mjs credentials
node backend/scripts/microsoft365-check.mjs mail
```

Режим `credentials` повторяет безопасную проверку секрета приложения и выводит только статус. Режим `mail` запускает listener только на loopback и сообщает локальный адрес входа. Открыть `http://localhost:8400/start` в браузере на том же компьютере и завершить Microsoft sign-in/MFA/consent. Запущенный на RnD listener без SSH tunnel не доступен через localhost браузера ноутбука.

Проверка использует tenant-specific Authorization Code + PKCE S256, одноразовый state и ограниченное время ожидания. В этом отдельном canary не запрашиваются `openid`, `offline_access`, ID token или refresh token: он проверяет Graph-доступ, не создаёт сессию Mike и не сохраняет подключение пользователя.

После возврата из Microsoft скрипт проверяет выданные `User.Read` и `Mail.Read`, выполняет `/me?$select=id`, затем `/me/messages?$top=1&$select=id,body`. Запрашивается не более одного письма; его тело не печатается и не сохраняется. Токен используется только в памяти процесса. Результат содержит HTTP-статусы, наличие разрешений, количество полученных сообщений и факт чтения body. Для пустого mailbox успешный запрос не объявляется доказательством чтения существующего письма.

## Как интерпретировать результат

| Результат | Что доказано / следующий шаг |
| --- | --- |
| Credentials HTTP 200 | Секрет приложения принят, но пользовательская почта ещё не проверена |
| Страница Microsoft требует вход/MFA | Пользователь завершает вход в браузере; пароль не передавать агенту или скрипту |
| `AADSTS50011` | Проверить Web redirect URI и точное написание callback path |
| Требуется admin approval | IT проверяет delegated consent и правила назначения приложения |
| Delegated token и `/me` успешны, mail отказан | Проверить `Mail.Read`, наличие Exchange Online mailbox и ограничения tenant; учитывать конкретный безопасный error code |
| Mail HTTP 200, mailbox пуст | Доступ к endpoint подтверждён; чтение body ещё не проверено |
| Mail HTTP 200, получено письмо с body | Delegated чтение одного письма подтверждено; это не проверка всей интеграции и изоляции пользователей |

В случае истечения времени или завершения процесса повторно запустить `mail` и открыть новую локальную ссылку. Старый state повторно использовать нельзя.

## Проверка самого диагностического скрипта

```bash
node --test backend/scripts/microsoft365-check.test.mjs
```

После успешного live mail canary продолжить [TODO](../TODO.md): production OAuth и encrypted token cache, привязка user/org/tenant, protected chats, source ACL, citations и дополнительные тесты. Этот скрипт не заменяет ни один из этих этапов.

## Microsoft documentation

- [Authorization Code + PKCE](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Redirect URI restrictions](https://learn.microsoft.com/en-us/entra/identity-platform/reply-url)
- [App-only tokens without roles](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow#controlling-tokens-without-the-roles-claim)
- [Get user](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0)
- [List messages](https://learn.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0)
