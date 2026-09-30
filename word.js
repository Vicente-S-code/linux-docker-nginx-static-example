
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, Header, Footer, PageBreak } = require('docx');
const fs = require('fs');

const FONT = "Arial";
const ACCENT = "1F4E79";
const b = { style: BorderStyle.SINGLE, size: 1, color: "BFBFBF" };
const borders = { top:b, bottom:b, left:b, right:b };

const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const BR = () => new Paragraph({ children: [new TextRun({ text: "", size: 4 })] });

function P(parts) {
  if (typeof parts === "string") parts = [{ text: parts }];
  return new Paragraph({
    spacing: { after: 160 },
    children: parts.map(p => new TextRun({ size: 22, font: FONT, ...p }))
  });
}

function Bullet(parts) {
  if (typeof parts === "string") parts = [{ text: parts }];
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: parts.map(p => new TextRun({ size: 22, font: FONT, ...p }))
  });
}

function Code(lines, dark = true) {
  return new Paragraph({
    shading: { fill: dark ? "1E1E1E" : "F5F5F5", type: ShadingType.CLEAR },
    border: { top: b, bottom: b, left: { style: BorderStyle.SINGLE, size: 12, color: dark ? "569CD6" : "BFBFBF" }, right: b },
    spacing: { before: 100, after: 180 },
    children: lines.flatMap((l, i) => {
      const isComment = l.trim().startsWith('#');
      const run = new TextRun({
        text: l, font: "Courier New", size: 19,
        color: dark ? (isComment ? "6A9955" : "DCDCAA") : "1F1F1F"
      });
      return i === 0 ? [run] : [new TextRun({ text: "", break: 1 }), run];
    })
  });
}

function AltBox(lines) {
  // Alternative command box (green tinted)
  return new Paragraph({
    shading: { fill: "E2EFDA", type: ShadingType.CLEAR },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "70AD47" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "70AD47" },
      left: { style: BorderStyle.SINGLE, size: 16, color: "70AD47" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "70AD47" }
    },
    spacing: { before: 100, after: 160 },
    children: lines.flatMap((l, i) => {
      const isComment = l.startsWith('#');
      const run = new TextRun({ text: l, font: "Courier New", size: 19, color: isComment ? "375623" : "1F6B0F" });
      return i === 0 ? [run] : [new TextRun({ text: "", break: 1 }), run];
    })
  });
}

function Tip(text) {
  return new Paragraph({
    shading: { fill: "E2EFDA", type: ShadingType.CLEAR },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "70AD47" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "70AD47" }, left: { style: BorderStyle.SINGLE, size: 16, color: "70AD47" }, right: { style: BorderStyle.SINGLE, size: 4, color: "70AD47" } },
    spacing: { before: 80, after: 160 },
    children: [new TextRun({ text: "✔  " + text, size: 20, color: "375623", font: FONT })]
  });
}

function Warn(text) {
  return new Paragraph({
    shading: { fill: "FFF2CC", type: ShadingType.CLEAR },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "BF8F00" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "BF8F00" }, left: { style: BorderStyle.SINGLE, size: 16, color: "BF8F00" }, right: { style: BorderStyle.SINGLE, size: 4, color: "BF8F00" } },
    spacing: { before: 80, after: 160 },
    children: [new TextRun({ text: "⚠  " + text, size: 20, color: "7F6000", font: FONT })]
  });
}

function Info(text) {
  return new Paragraph({
    shading: { fill: "DEEBF7", type: ShadingType.CLEAR },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6" }, left: { style: BorderStyle.SINGLE, size: 16, color: "2E75B6" }, right: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6" } },
    spacing: { before: 80, after: 160 },
    children: [new TextRun({ text: "ℹ  " + text, size: 20, color: "1F4E79", font: FONT })]
  });
}

function Screenshot(label) {
  return new Paragraph({
    shading: { fill: "FFF2CC", type: ShadingType.CLEAR },
    border: { top: { style: BorderStyle.DASHED, size: 6, color: "BF8F00" }, bottom: { style: BorderStyle.DASHED, size: 6, color: "BF8F00" }, left: { style: BorderStyle.DASHED, size: 6, color: "BF8F00" }, right: { style: BorderStyle.DASHED, size: 6, color: "BF8F00" } },
    spacing: { before: 120, after: 280 },
    children: [new TextRun({ text: "📷  CAPTURA: " + label, italics: true, size: 20, color: "7F6000", font: FONT })]
  });
}

function StepN(n, parts) {
  if (typeof parts === "string") parts = [{ text: parts }];
  return new Paragraph({
    spacing: { after: 100 },
    children: [
      new TextRun({ text: `PASO ${n}   `, bold: true, size: 22, color: ACCENT, font: FONT }),
      ...parts.map(p => new TextRun({ size: 22, font: FONT, ...p }))
    ]
  });
}

function cell(text, opts = {}) {
  return new TableCell({
    borders,
    width: { size: opts.w || 2000, type: WidthType.DXA },
    shading: opts.head ? { fill: ACCENT, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: !!opts.head, color: opts.head ? "FFFFFF" : "000000", size: 20, font: FONT })] })]
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: "FFFFFF" },
        paragraph: {
          spacing: { before: 400, after: 200 }, outlineLevel: 0,
          shading: { fill: ACCENT, type: ShadingType.CLEAR },
          indent: { left: 240, right: 240 }
        }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: ACCENT },
        paragraph: {
          spacing: { before: 300, after: 160 }, outlineLevel: 1,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "BDD7EE", space: 4 } }
        }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, italics: true, color: "2E75B6" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 }
      },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF", space: 4 } },
          children: [new TextRun({ text: "Guía Práctica Linux — Evaluación Investigativa TI3034", size: 16, color: "808080" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF", space: 4 } },
          children: [
            new TextRun({ text: "Carla — INACAP    |    Página ", size: 16, color: "808080" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "808080" }),
            new TextRun({ text: " de ", size: 16, color: "808080" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: "808080" })
          ]
        })]
      })
    },
    children: [

// ========= PORTADA =========
new Paragraph({ spacing: { before: 2000 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "GUÍA PRÁCTICA DE COMANDOS LINUX", bold: true, size: 44, color: ACCENT })] }),
new Paragraph({ spacing: { before: 120 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Evaluación Investigativa — Proyecto Integrador TI3034", size: 26 })] }),
new Paragraph({ spacing: { before: 80 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Fundamentos de Seguridad de la Información", size: 23, italics: true })] }),
new Paragraph({ spacing: { before: 900 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Estudiante: Carla", bold: true, size: 24 })] }),
new Paragraph({ spacing: { before: 80 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "INACAP — Junio 2026", size: 22 })] }),
new Paragraph({ children: [new PageBreak()] }),

// ========= INTROD =========
new Paragraph({
  shading: { fill: "DCE6F1", type: ShadingType.CLEAR },
  border: { top: b, bottom: b, left: { style: BorderStyle.SINGLE, size: 16, color: "2E75B6" }, right: b },
  spacing: { before: 80, after: 200 },
  children: [
    new TextRun({ text: "📌  CÓMO LEER ESTA GUÍA   ", bold: true, size: 21, color: ACCENT }),
    new TextRun({ text: "Los bloques oscuros son los comandos que ejecutas en la terminal. Los bloques verdes son alternativas equivalentes que también funcionan (provienen de las guías anteriores del curso). Cada 📷 indica el momento exacto en que debes tomar una captura de pantalla para tu informe.", size: 21, color: "1F4E79" })
  ]
}),
BR(),

// TABLA DE USUARIOS
H1("  Usuarios y grupos de este proyecto"),
BR(),
P("La configuración de este proyecto usa los siguientes usuarios y departamentos:"),
BR(),
new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2400, 2200, 2200, 2560],
  rows: [
    new TableRow({ children: [cell("Departamento", { w: 2400, head: true }), cell("Usuarios", { w: 2200, head: true }), cell("Grupo primario", { w: 2200, head: true }), cell("Notas", { w: 2560, head: true })] }),
    new TableRow({ children: [cell("DESARROLLO", { w: 2400 }), cell("carla, brian", { w: 2200 }), cell("DESARROLLO", { w: 2200 }), cell("Acceso solo a /empresa/desarrollo y /empresa/compartido", { w: 2560 })] }),
    new TableRow({ children: [cell("OPERACIONES", { w: 2400 }), cell("vicente, claudio", { w: 2200 }), cell("OPERACIONES", { w: 2200 }), cell("Acceso solo a /empresa/operaciones y /empresa/compartido", { w: 2560 })] }),
    new TableRow({ children: [cell("ADMINISTRACIÓN", { w: 2400 }), cell("bastian, lissette", { w: 2200 }), cell("ADMIN", { w: 2200 }), cell("bastian: acceso a TODAS las carpetas. lissette: solo /empresa/admin y /empresa/compartido", { w: 2560 })] }),
  ]
}),
BR(),
new Paragraph({ children: [new PageBreak()] }),

// ========= SECCION 0 =========
H1("  PASO 0 — Acceso como administrador (root)"),
BR(),
P("Antes de cualquier otra cosa, necesitas obtener permisos de administrador. Sin esto, useradd, groupadd, chown y chmod te darán error de permisos."),
BR(),
StepN(1, "Abre la terminal y escribe:"),
Code(["sudo -i"]),
P([{ text: "Qué hace: ", bold: true }, { text: "sudo -i simula un inicio de sesión completo como root. Carga todas las variables de entorno de root y te posiciona en /root. Es más estable que anteponer sudo a cada comando, porque desde aquí todos los comandos que escribas ya tienen permisos de administrador sin repetirlo." }]),
BR(),
Tip("Cómo saber que funcionó: el prompt cambia de usuario@host:~$ a root@host:~# — el símbolo # al final indica que ya eres root. A partir de aquí, TODOS los comandos siguientes van sin sudo."),
Screenshot("terminal mostrando el prompt cambiado a root@hostname:~#"),
BR(),

// ========= SECCION 1 =========
new Paragraph({ children: [new PageBreak()] }),
H1("  SECCIÓN 1 — Crear los grupos departamentales"),
BR(),
P("Los grupos son la base del control de acceso. Antes de crear usuarios, se crean los grupos para que al crear cada usuario pueda asignársele el grupo correcto desde el principio."),
BR(),
StepN(1, "Crear los cuatro grupos (uno por departamento más el grupo compartido):"),
Code(["groupadd DESARROLLO", "groupadd OPERACIONES", "groupadd ADMIN", "groupadd COMPARTIDO"]),
P([{ text: "Qué hace: ", bold: true }, { text: "groupadd crea un grupo nuevo en el sistema y le asigna un GID (Group ID) automáticamente. El GID queda registrado en el archivo /etc/group. Los nombres en mayúscula son solo una convención que facilita distinguir los grupos creados por nosotros de los del sistema (que suelen ir en minúscula)." }]),
BR(),
StepN(2, "Verificar que los 4 grupos se crearon correctamente:"),
Code(["tail -4 /etc/group"]),
P([{ text: "Qué hace: ", bold: true }, { text: "tail -4 muestra solo las últimas 4 líneas del archivo /etc/group, que son precisamente los 4 grupos que acabamos de crear, con su GID asignado." }]),
Screenshot("salida de tail -4 /etc/group mostrando DESARROLLO, OPERACIONES, ADMIN y COMPARTIDO con sus GID"),
BR(),

// ========= SECCION 2 =========
new Paragraph({ children: [new PageBreak()] }),
H1("  SECCIÓN 2 — Crear los usuarios"),
BR(),
P("Se crean 6 usuarios en total. Carla y Brian para Desarrollo, Vicente y Claudio para Operaciones, y Bastian y Lissette para Administración. Bastian es el supervisor y tendrá acceso a todos los departamentos."),
BR(),

H2("2.1 — Usuarios de DESARROLLO"),
StepN(1, "Crear el usuario carla con grupo primario DESARROLLO:"),
Code(["useradd -m -g DESARROLLO carla"]),
BR(),
StepN(2, "Crear el usuario brian con grupo primario DESARROLLO:"),
Code(["useradd -m -g DESARROLLO brian"]),
BR(),
P([{ text: "Qué hace cada parte: ", bold: true }]),
Bullet([{ text: "-m: ", bold: true }, { text: "crea automáticamente el directorio /home/carla y /home/brian, copiando los archivos base desde /etc/skel (que incluye el .bashrc por defecto)." }]),
Bullet([{ text: "-g DESARROLLO: ", bold: true }, { text: "define DESARROLLO como el grupo primario del usuario. El grupo primario es el que aparece en los archivos que el usuario crea." }]),
BR(),

H2("2.2 — Usuarios de OPERACIONES"),
StepN(3, "Crear el usuario vicente con grupo primario OPERACIONES:"),
Code(["useradd -m -g OPERACIONES vicente"]),
BR(),
StepN(4, "Crear el usuario claudio con grupo primario OPERACIONES:"),
Code(["useradd -m -g OPERACIONES claudio"]),
BR(),

H2("2.3 — Usuarios de ADMINISTRACIÓN"),
StepN(5, "Crear el usuario bastian con grupo primario ADMIN:"),
Code(["useradd -m -g ADMIN bastian"]),
BR(),
StepN(6, "Crear la usuaria lissette con grupo primario ADMIN:"),
Code(["useradd -m -g ADMIN lissette"]),
BR(),

H2("2.4 — Asignar contraseña a cada usuario"),
Warn("Sin contraseña asignada, la cuenta queda bloqueada y no se puede iniciar sesión con ella."),
BR(),
Code(["passwd carla", "passwd brian", "passwd vicente", "passwd claudio", "passwd bastian", "passwd lissette"]),
P("El sistema pedirá escribir la contraseña dos veces por cada usuario. No se ve nada mientras escribes, eso es normal."),
BR(),

H2("2.5 — Agregar usuarios al grupo COMPARTIDO"),
P("Todos los usuarios deben poder acceder a la carpeta colaborativa. Para eso se los agrega al grupo COMPARTIDO como grupo secundario (además de su grupo primario, que ya tienen):"),
BR(),
StepN(7, "Agregar a todos los usuarios al grupo COMPARTIDO:"),
Code(["usermod -aG COMPARTIDO carla", "usermod -aG COMPARTIDO brian", "usermod -aG COMPARTIDO vicente", "usermod -aG COMPARTIDO claudio", "usermod -aG COMPARTIDO lissette"]),
P([{ text: "Qué hace: ", bold: true }, { text: "usermod -aG agrega (append) al usuario como miembro secundario del grupo indicado, sin eliminar los grupos que ya tenía. Es fundamental el flag -a — sin él, se borrarían todos los grupos previos del usuario y quedaría solo en COMPARTIDO." }]),
BR(),

H2("2.6 — Hacer que bastian sea supervisor con acceso a todo"),
P("Bastian necesita acceso a las carpetas de Desarrollo, Operaciones y la carpeta compartida, además de su propia carpeta de Administración. Para eso se lo agrega a todos los grupos:"),
BR(),
StepN(8, "Asignar a bastian todos los grupos departamentales:"),
Code(["usermod -aG DESARROLLO,OPERACIONES,COMPARTIDO bastian"]),
P([{ text: "Qué hace: ", bold: true }, { text: "agrega a bastian como miembro secundario de los tres grupos a la vez. Ya tiene ADMIN como grupo primario (paso 5), así que después de este comando pertenece a los 4 grupos y podrá acceder a todas las carpetas del servidor." }]),
BR(),
Info("Los grupos secundarios se separan por coma sin espacios. Este comando es equivalente a ejecutar usermod -aG por separado para cada grupo, pero más eficiente."),
BR(),

H2("2.7 — Verificar que los usuarios quedaron bien"),
StepN(9, "Ver la lista de todos los usuarios creados:"),
Code(["tail -6 /etc/passwd"]),
Screenshot("salida de tail -6 /etc/passwd mostrando las 6 cuentas creadas"),
BR(),
StepN(10, [{ text: "Verificar los grupos de cada usuario con " }, { text: "id", font: "Courier New" }, { text: ":" }]),
Code(["id carla", "id vicente", "id bastian"]),
P([{ text: "Qué muestra: ", bold: true }, { text: "uid (identificador del usuario), gid (grupo primario) y groups (todos los grupos, primario y secundarios). Bastian debe aparecer en los 4 grupos: ADMIN, DESARROLLO, OPERACIONES y COMPARTIDO." }]),
Screenshot("salida de id carla, id vicente e id bastian — bastian debe mostrar los 4 grupos"),
BR(),
StepN(11, "Verificar también con el comando groups para mayor claridad:"),
Code(["groups bastian"]),
Tip("groups bastian debe mostrar: ADMIN DESARROLLO OPERACIONES COMPARTIDO. Si falta alguno, repite el usermod del paso 8."),
BR(),

// ========= SECCION 3 =========
new Paragraph({ children: [new PageBreak()] }),
H1("  SECCIÓN 3 — Crear la estructura de directorios"),
BR(),
StepN(1, "Crear la carpeta /empresa con todas sus subcarpetas de una sola vez:"),
Code(["mkdir -p /empresa/desarrollo /empresa/operaciones /empresa/admin /empresa/compartido"]),
P([{ text: "Qué hace: ", bold: true }, { text: "mkdir crea directorios. El flag -p (parents) crea automáticamente la carpeta padre /empresa si no existe, y luego las 4 subcarpetas juntas, todo en un solo comando. Sin -p tendríamos que crear /empresa primero y luego cada subcarpeta por separado." }]),
BR(),
StepN(2, "Verificar la estructura creada:"),
Code(["ls -la /empresa"]),
Screenshot("salida de ls -la /empresa mostrando las 4 subcarpetas recién creadas"),
BR(),

// ========= SECCION 4 =========
new Paragraph({ children: [new PageBreak()] }),
H1("  SECCIÓN 4 — Asignar grupos y permisos a las carpetas"),
BR(),
P("Este es el paso más importante de la implementación. Cada carpeta debe pertenecer al grupo de su departamento y tener los permisos correctos para que solo ese departamento pueda acceder a ella."),
BR(),
Warn("Asegúrate de seguir como root (prompt con #) antes de continuar."),
BR(),

H2("4.1 — Asignar el grupo propietario a cada carpeta"),
P("El comando chown (change owner) asigna quién es el dueño y qué grupo gestiona cada carpeta:"),
BR(),
StepN(1, "Asignar dueño y grupo a cada carpeta departamental:"),
Code([
  "chown root:DESARROLLO /empresa/desarrollo",
  "chown root:OPERACIONES /empresa/operaciones",
  "chown root:ADMIN /empresa/admin",
  "chown root:COMPARTIDO /empresa/compartido"
]),
P([{ text: "Qué hace: ", bold: true }, { text: "chown dueño:grupo carpeta asigna root como dueño (propietario) de la carpeta — esto evita que cualquier usuario pueda borrar o mover la carpeta raíz del departamento — y al grupo correspondiente como el grupo gestor, que es quien tendrá los permisos de acceso según el chmod siguiente." }]),
BR(),

H2("4.2 — Aplicar permisos a las carpetas departamentales"),
StepN(2, "Aplicar permisos 770 a cada carpeta departamental:"),
Code([
  "chmod 770 /empresa/desarrollo",
  "chmod 770 /empresa/operaciones",
  "chmod 770 /empresa/admin"
]),
P([{ text: "Qué hace 770: ", bold: true }, { text: "en notación octal, cada dígito representa permisos para Dueño / Grupo / Otros:" }]),
Bullet([{ text: "7 (rwx) ", bold: true }, { text: "para el dueño (root): control total — puede leer, escribir y entrar a la carpeta." }]),
Bullet([{ text: "7 (rwx) ", bold: true }, { text: "para el grupo del departamento: control total — los miembros del grupo pueden leer, escribir y entrar." }]),
Bullet([{ text: "0 (---) ", bold: true }, { text: "para \"otros\": sin ningún acceso. Así un usuario de Operaciones no puede ni siquiera ver qué hay en la carpeta de Administración." }]),
BR(),

H2("4.3 — Carpeta compartida con SGID"),
StepN(3, "Aplicar permisos con SGID a la carpeta compartida:"),
Code(["chmod 2770 /empresa/compartido"]),
P([{ text: "Qué hace el \"2\" inicial (SGID — Set Group ID): ", bold: true }, { text: "cualquier archivo o subcarpeta que se cree dentro de /empresa/compartido heredará automáticamente el grupo COMPARTIDO, sin importar cuál sea el grupo primario del usuario que lo creó. Sin esto, si Carla (Desarrollo) crea un archivo, quedaría con grupo DESARROLLO y Vicente (Operaciones) no podría editarlo." }]),
BR(),
P("También se puede aplicar permisos 770 más SGID de esta forma, que es equivalente:"),
AltBox([
  "# Alternativa con dos comandos separados (misma función):",
  "chmod 770 /empresa/compartido",
  "chmod g+s /empresa/compartido"
]),
P([{ text: "chmod g+s ", font: "Courier New", bold: true }, { text: "activa el bit SGID directamente sobre el grupo de la carpeta, sin cambiar los permisos rwx ya configurados. Ambas formas producen exactamente el mismo resultado: los permisos quedan en 2770 (drwxrws---)." }]),
BR(),

H2("4.4 — Verificar todo junto"),
StepN(4, "Ver los permisos finales de las 4 carpetas:"),
Code(["ls -la /empresa"]),
P("Deberías ver algo así:"),
Code([
  "drwxrwx---  root  DESARROLLO   desarrollo/",
  "drwxrwx---  root  OPERACIONES  operaciones/",
  "drwxrwx---  root  ADMIN        admin/",
  "drwxrws---  root  COMPARTIDO   compartido/"
], false),
P([{ text: "Fíjate en la \"s\" ", bold: true }, { text: "en lugar de la \"x\" del grupo en la carpeta compartido — esa s confirma visualmente que el SGID está activo. Esa captura es clave para el informe." }]),
Screenshot("ls -la /empresa mostrando los 4 directorios con sus grupos y permisos — observar la 's' en compartido"),
BR(),
StepN(5, "Probar que las restricciones funcionan — cambiar al usuario carla e intentar acceder a otra carpeta:"),
Code(["su - carla"]),
Code(["cd /empresa/desarrollo", "# debe funcionar (carla pertenece a DESARROLLO)", "cd /empresa/operaciones", "# debe dar Permission denied"]),
P("Volver a root:"),
Code(["exit"]),
Screenshot("terminal mostrando que carla puede entrar a /empresa/desarrollo pero recibe Permission denied al intentar /empresa/operaciones"),
BR(),

// ========= SECCION 5 =========
new Paragraph({ children: [new PageBreak()] }),
H1("  SECCIÓN 5 — Personalizar el .bashrc de cada usuario"),
BR(),
P("El archivo .bashrc se carga cada vez que el usuario abre una terminal. Aquí se personaliza el prompt (PS1), los alias y los mensajes de bienvenida. Se edita con nano, igual que en las guías anteriores."),
Warn("Edita siempre al FINAL del .bashrc, sin borrar el contenido que ya existe. El contenido original del sistema es necesario para que la terminal funcione correctamente."),
BR(),

H2("5.1 — .bashrc de carla (Desarrollo)"),
StepN(1, "Abrir el .bashrc de carla:"),
Code(["nano /home/carla/.bashrc"]),
P("Bájate hasta el final del archivo (Ctrl+End) y agrega estas líneas:"),
Code([
  "# ============================================",
  "# Configuracion personalizada - DESARROLLO",
  "# ============================================",
  "",
  "# Mensaje de bienvenida",
  "echo \"Bienvenido al entorno de Desarrollo, $(whoami)\"",
  "",
  "# Prompt: usuario@hostname:directorio$",
  "PS1='\\u@\\h:\\w$ '",
  "",
  "# Historial amplio para sesiones de desarrollo largas",
  "HISTSIZE=5000",
  "",
  "# Aliases utiles",
  "alias ll='ls -la --color=auto'",
  "alias gst='git status'",
  "alias dps='docker ps'"
]),
P("Guardar: Ctrl+O → Enter. Salir: Ctrl+X."),
P("Repetir exactamente lo mismo para brian:"),
Code(["nano /home/brian/.bashrc"]),
Screenshot("nano mostrando el .bashrc de carla con el bloque personalizado agregado al final"),
BR(),

H2("5.2 — .bashrc de vicente (Operaciones)"),
StepN(2, "Abrir el .bashrc de vicente:"),
Code(["nano /home/vicente/.bashrc"]),
P("Agregar al final:"),
Code([
  "# ============================================",
  "# Configuracion personalizada - OPERACIONES",
  "# ============================================",
  "",
  "# Funcion que muestra estado del sistema al abrir terminal",
  "resumen_usuario() {",
  "  echo \"=== ESTADO DE SESION PARA $(whoami) ===\"",
  "  echo \"Host: $(hostname) | Directorio: $(pwd)\"",
  "  uptime",
  "}",
  "resumen_usuario",
  "",
  "# Prompt con hora: [HH:MM:SS] usuario [directorio]",
  "PS1='[\\t] \\u [\\w] '",
  "",
  "# Aliases de monitoreo",
  "alias mem='free -m'",
  "alias disk='df -h'",
  "alias cpu='top -b -n1 | head -n 10'"
]),
P("Guardar y repetir para claudio:"),
Code(["nano /home/claudio/.bashrc"]),
Screenshot("nano mostrando el .bashrc de vicente con el bloque de operaciones agregado al final"),
BR(),

H2("5.3 — .bashrc de bastian (Administración)"),
StepN(3, "Abrir el .bashrc de bastian:"),
Code(["nano /home/bastian/.bashrc"]),
P("Agregar al final:"),
Code([
  "# ============================================",
  "# Configuracion personalizada - ADMINISTRACION",
  "# ============================================",
  "",
  "# Prompt en rojo que indica claramente sesion de administrador",
  "PS1='[ADMIN]-\\u@\\h:\\w$ '",
  "",
  "# Historial extendido para auditoria y trazabilidad",
  "export HISTSIZE=10000",
  "export HISTFILESIZE=20000",
  "export HISTTIMEFORMAT='%F %T '",
  "",
  "# Aliases administrativos",
  "alias usuarios='cut -d: -f1 /etc/passwd'",
  "alias grupos='cat /etc/group'",
  "alias permisos_empresa='ls -ld /empresa /empresa/*'",
  "",
  "echo \"ALERTA: Sesion de ADMINISTRADOR iniciada. Todas las acciones son auditadas.\""
]),
P("Guardar y repetir para lissette:"),
Code(["nano /home/lissette/.bashrc"]),
Screenshot("nano mostrando el .bashrc de bastian con el bloque de administración"),
BR(),

H2("5.4 — Verificar que los cambios funcionan"),
StepN(4, "Revisar el contenido completo de cada .bashrc:"),
Code(["cat /home/carla/.bashrc", "cat /home/vicente/.bashrc", "cat /home/bastian/.bashrc"]),
P("Esto imprime el contenido del archivo en pantalla para confirmar que no quedaron errores de tipeo."),
BR(),
StepN(5, "Probar el prompt real de cada usuario abriendo una sesión:"),
Code(["su - carla"]),
P("Debes ver el mensaje de bienvenida y el prompt: carla@hostname:~$"),
Code(["exit"]),
Code(["su - vicente"]),
P("Debes ver el resumen del sistema y el prompt con hora: [14:32:10] vicente [~]"),
Code(["exit"]),
Code(["su - bastian"]),
P("Debes ver la alerta de administrador y el prompt: [ADMIN]-bastian@hostname:~$"),
Code(["exit"]),
Screenshot("terminal mostrando los tres prompts en acción: carla, vicente con hora, bastian con [ADMIN]"),
BR(),

// ========= SECCION 6 =========
new Paragraph({ children: [new PageBreak()] }),
H1("  SECCIÓN 6 — Políticas de contraseña con chage"),
BR(),
P("El comando chage configura cuándo vence la contraseña de cada usuario y cada cuánto tiempo debe cambiarse. Estos parámetros se almacenan directamente en /etc/shadow."),
BR(),
P([{ text: "Qué significa cada flag: ", bold: true }]),
Bullet([{ text: "-m ", bold: true, font: "Courier New" }, { text: "(mínimo): días mínimos que deben pasar entre un cambio de contraseña y el siguiente." }]),
Bullet([{ text: "-M ", bold: true, font: "Courier New" }, { text: "(máximo): días antes de que la contraseña expire y se obligue a renovarla." }]),
Bullet([{ text: "-W ", bold: true, font: "Courier New" }, { text: "(warning): días antes de la expiración en que el sistema empieza a avisarle al usuario." }]),
Bullet([{ text: "-I ", bold: true, font: "Courier New" }, { text: "(inactive): días de gracia tras la expiración antes de que la cuenta se bloquee automáticamente por inactividad." }]),
BR(),

H2("6.1 — Desarrollo (carla y brian) — política estándar"),
Code([
  "chage -m 5 -M 90 -W 10 -I 5 carla",
  "chage -m 5 -M 90 -W 10 -I 5 brian"
]),
P("Mínimo 5 días entre cambios, máximo 90 días de validez, aviso 10 días antes, gracia de 5 días tras vencer."),
BR(),

H2("6.2 — Operaciones (vicente y claudio) — política más relajada"),
Code([
  "chage -m 10 -M 180 -W 25 -I 30 vicente",
  "chage -m 10 -M 180 -W 25 -I 30 claudio"
]),
P("Contraseña válida por 180 días, aviso 25 días antes, gracia de 30 días. Política más permisiva dado el menor riesgo de las cuentas de monitoreo."),
BR(),

H2("6.3 — Administración (bastian y lissette) — política estricta"),
Code([
  "chage -m 1 -M 30 -W 7 -I 3 bastian",
  "chage -m 1 -M 30 -W 7 -I 3 lissette"
]),
P("La más estricta: contraseña vence cada 30 días, apenas 3 días de gracia. Justificado porque bastian tiene acceso a todos los departamentos."),
BR(),

H2("6.4 — Verificar las políticas configuradas"),
StepN(1, "Ver la política de un usuario por departamento:"),
Code(["chage -l carla", "chage -l vicente", "chage -l bastian"]),
P([{ text: "Qué muestra: ", bold: true }, { text: "chage -l (list) imprime en formato legible las fechas de vencimiento, los días mínimos/máximos, los días de aviso y la fecha de inactividad. Esto es lo que debes mostrar en el informe." }]),
Screenshot("salida de chage -l carla, chage -l vicente y chage -l bastian mostrando las políticas de cada departamento"),
BR(),

// ========= SECCION 7 =========
new Paragraph({ children: [new PageBreak()] }),
H1("  SECCIÓN 7 — Scripts bash (Bloque 4)"),
BR(),
P("Se crean dos scripts con nano. El primero automatiza toda la configuración del servidor. El segundo audita que todo esté correctamente implementado."),
BR(),

H2("7.1 — Script de configuración automática (init_estructura.sh)"),
StepN(1, "Crear el archivo del script:"),
Code(["nano /root/init_estructura.sh"]),
BR(),
P("Escribe el siguiente contenido completo:"),
Code([
  "#!/bin/bash",
  "# init_estructura.sh",
  "# Crea grupos, directorios y permisos automaticamente",
  "# Ejecutar como root: bash /root/init_estructura.sh",
  "",
  "# Verificar que se ejecuta como root",
  "if [ \"$(id -u)\" -ne 0 ]; then",
  "  echo \"Error: este script debe ejecutarse como root.\"",
  "  exit 1",
  "fi",
  "",
  "echo \"=== Creando grupos ===\"",
  "groupadd -f DESARROLLO",
  "groupadd -f OPERACIONES",
  "groupadd -f ADMIN",
  "groupadd -f COMPARTIDO",
  "",
  "echo \"=== Creando directorios ===\"",
  "mkdir -p /empresa/desarrollo /empresa/operaciones /empresa/admin /empresa/compartido",
  "",
  "echo \"=== Asignando grupos ===\"",
  "chown root:DESARROLLO /empresa/desarrollo",
  "chown root:OPERACIONES /empresa/operaciones",
  "chown root:ADMIN /empresa/admin",
  "chown root:COMPARTIDO /empresa/compartido",
  "",
  "echo \"=== Aplicando permisos ===\"",
  "chmod 770 /empresa/desarrollo /empresa/operaciones /empresa/admin",
  "chmod 2770 /empresa/compartido",
  "",
  "echo \"=== Configuracion completada ===\"",
  "ls -la /empresa"
]),
BR(),
StepN(2, "Guardar y salir: Ctrl+O → Enter → Ctrl+X"),
BR(),
StepN(3, "Dar permiso de ejecución al script:"),
Code(["chmod +x /root/init_estructura.sh"]),
P([{ text: "Qué hace chmod +x: ", bold: true }, { text: "agrega el bit de ejecución al archivo. Sin esto el sistema no puede correrlo como programa, aunque tenga el #!/bin/bash al inicio." }]),
BR(),
P("También se puede hacer así (equivalente, como vimos en la Guía 2):"),
AltBox([
  "# Alternativa — solo dar permiso de ejecucion al dueno:",
  "chmod u+x /root/init_estructura.sh"
]),
BR(),
StepN(4, "Ejecutar el script:"),
Code(["bash /root/init_estructura.sh"]),
Screenshot("terminal mostrando la ejecución completa de init_estructura.sh sin errores"),
BR(),

H2("7.2 — Script de auditoría (auditoria_permisos.sh)"),
StepN(5, "Crear el archivo de auditoría:"),
Code(["nano /root/auditoria_permisos.sh"]),
BR(),
P("Escribe el siguiente contenido:"),
Code([
  "#!/bin/bash",
  "# auditoria_permisos.sh",
  "# Verifica que la configuracion de /empresa sea correcta",
  "",
  "echo \"--- AUDITORIA DE PERMISOS /EMPRESA ---\"",
  "ls -la /empresa",
  "",
  "echo \"\"",
  "echo \"--- VERIFICACION DE SGID EN COMPARTIDO ---\"",
  "ls -ld /empresa/compartido | grep \"s\" && echo \"SGID: OK\" || echo \"SGID: ERROR\"",
  "",
  "echo \"\"",
  "echo \"--- GRUPOS DEL SUPERVISOR (bastian) ---\"",
  "id bastian",
  "groups bastian",
  "",
  "echo \"\"",
  "echo \"--- POLITICAS DE CONTRASENA ---\"",
  "for u in carla vicente bastian; do",
  "  echo \"--- Usuario: $u ---\"",
  "  chage -l \"$u\" | grep -E \"Password expires|Maximum|Minimum\"",
  "done"
]),
BR(),
StepN(6, "Guardar y dar permiso de ejecución:"),
Code(["chmod +x /root/auditoria_permisos.sh"]),
BR(),
StepN(7, "Ejecutar la auditoría:"),
Code(["bash /root/auditoria_permisos.sh"]),
P("Este script genera automáticamente un reporte con los permisos actuales, el estado del SGID, los grupos de bastian y las políticas de contraseña — todo en un solo comando."),
Screenshot("salida completa de auditoria_permisos.sh mostrando permisos, SGID OK, grupos de bastian y políticas"),
BR(),

// ========= SECCION 8 =========
new Paragraph({ children: [new PageBreak()] }),
H1("  SECCIÓN 8 — Verificación final"),
BR(),
P("Antes de cerrar la terminal, confirma que todo quedó correctamente con estos comandos de repaso:"),
BR(),
Code([
  "# 1. Ver todos los usuarios creados",
  "tail -6 /etc/passwd",
  "",
  "# 2. Ver grupos y sus miembros",
  "tail -4 /etc/group",
  "",
  "# 3. Permisos de /empresa",
  "ls -la /empresa",
  "",
  "# 4. Grupos del supervisor bastian",
  "id bastian",
  "",
  "# 5. Politicas de los 3 departamentos",
  "chage -l carla",
  "chage -l vicente",
  "chage -l bastian"
]),
Screenshot("captura final mostrando todos los comandos de verificación ejecutados correctamente"),
BR(),
Tip("Si todo está correcto: existen los 6 usuarios, los 4 grupos, las 4 carpetas tienen sus grupos y permisos, la 's' aparece en /empresa/compartido, bastian está en los 4 grupos y chage muestra las políticas correctas para cada departamento."),
BR(),

// ========= RESUMEN COMANDOS =========
new Paragraph({ children: [new PageBreak()] }),
H1("  RESUMEN DE TODOS LOS COMANDOS USADOS"),
BR(),
new Table({
  width: { size: 10080, type: WidthType.DXA },
  columnWidths: [3200, 3880, 3000],
  rows: [
    new TableRow({ children: [
      new TableCell({ borders, shading:{fill:ACCENT,type:ShadingType.CLEAR}, margins:{top:80,bottom:80,left:120,right:120}, width:{size:3200,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:"Comando",bold:true,color:"FFFFFF",size:20})]})] }),
      new TableCell({ borders, shading:{fill:ACCENT,type:ShadingType.CLEAR}, margins:{top:80,bottom:80,left:120,right:120}, width:{size:3880,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:"¿Qué hace?",bold:true,color:"FFFFFF",size:20})]})] }),
      new TableCell({ borders, shading:{fill:ACCENT,type:ShadingType.CLEAR}, margins:{top:80,bottom:80,left:120,right:120}, width:{size:3000,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:"Alternativa vista en guías",bold:true,color:"FFFFFF",size:20})]})] }),
    ]}),
    ...[
      ["sudo -i","Acceder como root con entorno completo","sudo su -"],
      ["groupadd NOMBRE","Crear un grupo nuevo","(mismo en Guía 3)"],
      ["useradd -m -g GRUPO usuario","Crear usuario con home y grupo primario","useradd -m -s /bin/bash -g GRUPO usuario"],
      ["passwd usuario","Asignar contraseña a un usuario","(mismo en Guías 2 y 3)"],
      ["usermod -aG GRUPOS usuario","Agregar usuario a grupos secundarios sin borrar los anteriores","(nuevo en esta guía)"],
      ["mkdir -p /ruta/completa","Crear directorios con sus padres","(mismo en Guía 2)"],
      ["chown root:GRUPO carpeta","Asignar dueño y grupo a una carpeta","(nuevo en esta guía — alternativa: crear la dir como ese usuario)"],
      ["chmod 770 carpeta","Permisos rwxrwx--- — solo dueño y grupo","(mismo en Guía 2)"],
      ["chmod 2770 carpeta","770 más SGID en un solo comando","chmod 770 + chmod g+s (Guía 3)"],
      ["chmod +x archivo","Dar permiso ejecutable (a todos)","chmod u+x archivo (Guía 2)"],
      ["chage -m -M -W -I usuario","Configurar política de contraseña","(mismo en Guía 3)"],
      ["chage -l usuario","Ver política de contraseña activa","(mismo en Guía 3)"],
      ["id usuario","Ver uid, grupo primario y todos los grupos","(mismo en Guía 3)"],
      ["groups usuario","Ver solo los grupos de un usuario","(nuevo en esta guía)"],
      ["tail -N archivo","Ver las últimas N líneas de un archivo","cat archivo (Guías 2 y 3)"],
      ["su - usuario","Cambiar a otro usuario con su entorno","(mismo en Guías 2 y 3)"],
      ["nano archivo","Editar un archivo de texto","(mismo en Guías 2 y 3)"],
    ].map(([cmd, desc, alt]) => new TableRow({ children: [
      new TableCell({ borders, margins:{top:80,bottom:80,left:120,right:120}, width:{size:3200,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:cmd,font:"Courier New",size:18,color:"C55A11"})]})] }),
      new TableCell({ borders, margins:{top:80,bottom:80,left:120,right:120}, width:{size:3880,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:desc,size:19})]})] }),
      new TableCell({ borders, margins:{top:80,bottom:80,left:120,right:120}, width:{size:3000,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun({text:alt,size:19,italics:true,color:"375623"})]})] }),
    ]}))
  ]
}),
BR(),
BR(),

// CHECKLIST
H1("  CHECKLIST DE CAPTURAS DE PANTALLA"),
BR(),
P("Según lo que pide la evaluación, asegúrate de tener captura de:"),
BR(),
Bullet("tail -4 /etc/group → los 4 grupos creados"),
Bullet("tail -6 /etc/passwd → los 6 usuarios creados"),
Bullet("id carla, id vicente, id bastian → verificar grupos (bastian debe tener los 4)"),
Bullet("groups bastian → confirmar acceso total del supervisor"),
Bullet("ls -la /empresa → estructura de carpetas con permisos y la 's' del SGID en compartido"),
Bullet("su - carla intentando entrar a /empresa/operaciones → Permission denied"),
Bullet("cat de cada .bashrc personalizado o captura del prompt con su - usuario"),
Bullet("chage -l carla, chage -l vicente, chage -l bastian → políticas por departamento"),
Bullet("Salida completa del script auditoria_permisos.sh ejecutándose"),
    ]
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync("C:\Users\sistemas\Desktop\word2.docx", b);
  console.log("done");
});
