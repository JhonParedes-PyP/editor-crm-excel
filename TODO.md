# TODO: Mejoras App Editor Excel → CRM
Estado: En progreso | Aprobado por usuario

## Pasos del Plan (marcar con [x] al completar)

### 1. Setup Inicial [x]
- [x] Actualizar requirements.txt con nuevas deps
- [x] Crear config.py (columnas CRM)
- [x] Crear utils.py (funciones reutilizables: mapeo, validaciones)

### 2. Mejoras UI/UX [x]
- [x] Soporte multi-hoja por archivo
- [x] Edición inline en preview (st.data_editor o aggrid)
- [ ] Drag&drop files
- [ ] Dashboard stats (total deuda, % mora, top deudores)
- [ ] Filtros/búsqueda por COD_CREDITO/DNI
- [ ] Temas dark/light

### 3. Validaciones & Reports [ ]
- [ ] Regex DNI/RUC
- [ ] Cálculos derivados (RANGO_DIAS_MORA)
- [ ] Reporte PDF summary
- [ ] Logs duplicados exportables

### 4. Export & Performance [ ]
- [ ] Export CSV/JSON
- [ ] Paginación previews grandes
- [ ] Límites tamaño archivo

### 5. Testing & Docs [ ]
- [ ] Crear tests/test_app.py
- [ ] README.md completo
- [ ] Probar: `streamlit run app.py`
- [ ] Marcar TODO completo y cleanup

**Próximo paso: Dashboard stats y filtros**

