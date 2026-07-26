# TODO - Sincronizar categorías y agregar teléfono ✅

## 1. Código.gs - Backend ✅
- [x] Agregar columna 'Telefono' al header en `getUsersSheet()`
- [x] Guardar `telefono` en `loginOrRegister()` al registrar nuevo usuario
- [x] Actualizar/guardar `telefono` también en `submitVote()`
- [x] Incluir `telefono` en los resultados de `getWinner()` y `getResponses()`
- [x] **Corregir** columna de lectura de contraseña admin: columna 4 → columna 6
- [x] **Corregir** `getConfigSheet()` para que SIEMPRE escriba 'baby2026' si está vacía
- [x] Agregar función `verifyAdmin()` para validar contraseña desde admin.html

## 2. index.html - Frontend de usuarios ✅
- [x] Agregar campo "📞 Número de contacto" en la pantalla de login
- [x] Darle el mismo estilo CSS que usuario y contraseña (agregado `input[type="tel"]`)
- [x] Enviar `telefono` en la petición `loginOrRegister`
- [x] Enviar `telefono` en la petición `submitVote`

## 3. admin.html - Panel de administración ✅
- [x] Agregar selector "Tipo de parto" (Parto/Cesárea)
- [x] Agregar selector "Semana de nacimiento" (Semanas 35-40)
- [x] Corregir opciones de "Hora" para que coincidan con index.html (4 rangos)
- [x] Enviar las 5 categorías completas en `setAnswer`
- [x] Actualizar tabla para mostrar Parto, Semana y Teléfono
- [x] Cambiar métrica de "/ 3" a "/ 5"
- [x] Usar `verifyAdmin` para login seguro

