---
name: GitHub preference
description: Пользователь хочет, чтобы все файлы автоматически пушились на GitHub
---

**Правило:** после каждого завершённого куска работы пушить файлы на GitHub.

**Репозиторий:** `https://github.com/ScarFace11/contract-prolongation`  
**Ветка:** `main`  
**Remote `origin` уже настроен** в проекте.

**Как применять:** после создания или изменения файлов — `git add`, `git commit`, затем `gitPush({ branch: "main", provider: "github" })` через CodeExecution.

**Why:** пользователь явно попросил хранить все файлы проекта в этом репозитории.
