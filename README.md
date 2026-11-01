# Спрашиватель — запуск в Docker

## Первый запуск
1. Установите Docker (на Windows — Docker Desktop).
2. Скопируйте `.env.example` в `.env` и впишите `MAX_BOT_TOKEN`.
3. В папке проекта выполните:

       docker compose up -d --build

4. Логи: `docker compose logs -f`

## Управление
- Остановить: `docker compose down` (данные сохраняются)
- Перезапустить: `docker compose restart`
- Обновить бота (заменили `bot.py`): `docker compose up -d --build`

## Где хранятся данные
Опросы и ответы лежат в томе Docker `sprashivatel_data` (файл `/data/sprashivatel.db`), суточные копии — в `/data/backups`.
Том не удаляется при пересборке и `docker compose down`. **Не запускайте** `docker compose down -v` — флаг `-v` удалит данные.

## Резервная копия и перенос
Скопировать базу на компьютер:

    docker compose cp sprashivatel:/data/sprashivatel.db ./sprashivatel_backup.db

Положить базу обратно (бот остановлен):

    docker compose stop
    docker compose cp ./sprashivatel_backup.db sprashivatel:/data/sprashivatel.db
    docker compose start

Если у Вас уже есть `sprashivatel.db` от запуска без Docker — перенесите её так же (команда выше).

## Важно
- Запускайте **одну** копию бота с одним токеном. Если бот ещё запущен на компьютере без Docker, остановите его.
- Бот в чате должен быть администратором (см. справку в боте).
