%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%  plantilla-uniandes.cls
%%  Reproduccion en LaTeX de "Plantilla_Proyecto_de_Grado.docx"
%%  Universidad de los Andes - Ing. de Sistemas y Computacion
%%
%%  Todos los valores de esta clase fueron extraidos directamente del XML
%%  interno del archivo .docx (styles.xml / document.xml / settings.xml),
%%  de modo que el resultado sea metricamente identico al original.
%%
%%  Compilar con XeLaTeX o LuaLaTeX (recomendado, usa Arial real).
%%  pdfLaTeX tambien funciona: usa Helvetica/Nimbus Sans, clon metrico de Arial.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{plantilla-uniandes}[2026/07/29 v1.0 Plantilla Proyecto de Grado Uniandes]

%%-----------------------------------------------------------------------------
%% Clase base
%%   Word: pgSz 12240x15840 twips = carta (letter), orientacion vertical.
%%   Word: mirrorMargins (margenes simetricos) -> twoside
%%   Word: cada capitulo abre en pagina impar (sectPr type="oddPage") -> openright
%%   Word: Normal = Arial 11 pt (sz 22 half-points)
%%-----------------------------------------------------------------------------
\DeclareOption*{\PassOptionsToClass{\CurrentOption}{book}}
\ProcessOptions\relax
\LoadClass[11pt,letterpaper,twoside,openright]{book}

\RequirePackage{iftex}

%%-----------------------------------------------------------------------------
%% Codificacion, idioma y tipografia
%%   styles.xml -> <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/> en el estilo Normal
%%-----------------------------------------------------------------------------
\ifPDFTeX
  \RequirePackage[T1]{fontenc}
  \RequirePackage[utf8]{inputenc}
  \RequirePackage{helvet}          % Nimbus Sans: clon metrico de Helvetica/Arial
  \renewcommand{\familydefault}{\sfdefault}
\else
  \RequirePackage{fontspec}
  \IfFontExistsTF{Arial}{%
    \setmainfont{Arial}%
  }{%
    \IfFontExistsTF{Liberation Sans}{%
      \setmainfont{Liberation Sans}%   metricamente identica a Arial
    }{%
      \setmainfont{TeX Gyre Heros}%    clon de Helvetica
    }%
  }
\fi

% Se carga el modulo de espanol de babel si esta instalado (division silabica
% en espanol); en caso contrario se continua sin el, ya que todos los rotulos
% se definen explicitamente mas abajo.
\IfFileExists{spanish.ldf}{%
  \RequirePackage[spanish,es-tabla,es-noquoting]{babel}%
}{%
  \RequirePackage[english]{babel}%
}
\RequirePackage{ragged2e}

%%-----------------------------------------------------------------------------
%% Geometria de pagina
%%   Word pgMar: top=1440 tw (2,54 cm)  bottom=1440 tw (2,54 cm)
%%               left(interior)=2041 tw (3,60 cm)  right(exterior)=1440 tw (2,54 cm)
%%               header=709 tw (1,25 cm desde el borde superior)
%%               footer=709 tw (1,25 cm desde el borde inferior)
%%-----------------------------------------------------------------------------
\RequirePackage[%
  letterpaper,
  twoside,
  top=2.54cm,
  bottom=2.54cm,
  inner=3.60cm,
  outer=2.54cm,
  headheight=14pt,
  headsep=0.868cm,     % situa el texto del encabezado a 1,25 cm del borde
                       % superior, igual que <w:pgMar w:header="709"/> en Word
  footskip=1.25cm
]{geometry}

%%-----------------------------------------------------------------------------
%% Interlineado
%%   El texto del documento indica: "interlineado a 1.5 lineas".
%%   El estilo de cuerpo (Prrafodelista) usa line=360 twips = 1,5 lineas.
%%   Word calcula 1,5 lineas de Arial 11 pt = 18,98 pt de linea base;
%%   \baselineskip de LaTeX a 11 pt = 13,6 pt  ->  factor 1,4.
%%-----------------------------------------------------------------------------
\RequirePackage{setspace}
\newcommand{\interlineadoTesis}{\setstretch{1.4}}
% OJO: \setstretch de setspace llama a \@currsize y anula un \fontsize previo.
% \linespread+\selectfont conserva el tamano ya seleccionado.
\newcommand{\interlineadoSencillo}{\linespread{1.0}\selectfont}
\AtBeginDocument{\interlineadoTesis}

%%-----------------------------------------------------------------------------
%% Parrafos
%%   Estilo de cuerpo: jc="both" (justificado), sin sangria,
%%   spacing after=0 y separacion entre parrafos con una linea en blanco
%%   ("despues de punto aparte se dejan dos interlineas").
%%-----------------------------------------------------------------------------
\setlength{\parindent}{0pt}
% \parskip debe valer una linea completa, pero \baselineskip solo tiene su
% valor definitivo cuando ya se aplicaron \normalsize y \setstretch; por eso
% se fija al comenzar el documento.
\AtBeginDocument{%
  \normalsize
  \setlength{\parskip}{\baselineskip}%
  \justifying
}

%%-----------------------------------------------------------------------------
%% Titulos
%%   Ttulo1  (Titulo Primer Nivel):  Arial 20 pt negrilla, antes 100 pt, despues 24 pt
%%   Ttulo2  (Titulo Segundo nivel): Arial 16 pt negrilla, antes  20 pt, despues 10 pt
%%   Ttulo3  (Titulo Tercer nivel):  Arial 14 pt negrilla, antes  20 pt, despues 10 pt
%%   Ttulo4  (Titulo Cuarto nivel):  Arial 11 pt normal,   antes  20 pt, despues 10 pt
%%   Subttulo (Titulo Anexos):       Arial 21 pt negrilla, antes 100 pt, despues 24 pt
%%   Ttulospreliminares2:            Arial 20 pt negrilla, antes 100 pt
%%   Ttulospreliminares:             Arial 18 pt negrilla, antes 100 pt, despues 10 pt
%%   Solo se numera hasta el tercer nivel.
%%-----------------------------------------------------------------------------
\RequirePackage{titlesec}

\setcounter{secnumdepth}{2}   % chapter(0), section(1), subsection(2)
\setcounter{tocdepth}{2}

% LaTeX inserta \parskip antes y despues de cada titulo; para que el espacio
% neto coincida con el de Word hay que descontarlo. Las longitudes se calculan
% al comenzar el documento, cuando \parskip ya tiene su valor definitivo.
\newlength{\espAnteCapitulo}   \newlength{\espTrasCapitulo}
\newlength{\espAnteSeccion}    \newlength{\espTrasSeccion}
\AtBeginDocument{%
  \setlength{\espAnteCapitulo}{\dimexpr100pt-\parskip\relax}%
  \setlength{\espTrasCapitulo}{\dimexpr 24pt-\parskip\relax}%
  \setlength{\espAnteSeccion} {\dimexpr 20pt-\parskip\relax}%
  \setlength{\espTrasSeccion} {\dimexpr 10pt-\parskip\relax}%
  \ifdim\espTrasCapitulo<0pt \setlength{\espTrasCapitulo}{0pt}\fi
  \ifdim\espAnteSeccion<0pt  \setlength{\espAnteSeccion}{0pt}\fi
  \ifdim\espTrasSeccion<0pt  \setlength{\espTrasSeccion}{0pt}\fi
}

\titleformat{\chapter}[hang]
  {\normalfont\interlineadoSencillo\bfseries\fontsize{20}{20}\selectfont\raggedright}
  {\thechapter.}{0.5em}{}
\titlespacing*{\chapter}{0pt}{\espAnteCapitulo}{\espTrasCapitulo}

\titleformat{\section}[hang]
  {\normalfont\interlineadoSencillo\bfseries\fontsize{16}{16}\selectfont\raggedright}
  {\thesection}{0.5em}{}
\titlespacing*{\section}{0pt}{\espAnteSeccion}{\espTrasSeccion}

\titleformat{\subsection}[hang]
  {\normalfont\interlineadoSencillo\bfseries\fontsize{14}{14}\selectfont\raggedright}
  {\thesubsection}{0.5em}{}
\titlespacing*{\subsection}{0pt}{\espAnteSeccion}{\espTrasSeccion}

\titleformat{\subsubsection}[hang]
  {\normalfont\interlineadoSencillo\fontsize{11}{13}\selectfont\raggedright}
  {}{0pt}{}
\titlespacing*{\subsubsection}{0pt}{\espAnteSeccion}{\espTrasSeccion}

% Numeracion sin punto final en niveles 2 y 3 (Word: "1.1", "1.1.1")
\renewcommand{\thesection}{\thechapter.\arabic{section}}
\renewcommand{\thesubsection}{\thesection.\arabic{subsection}}

%%-----------------------------------------------------------------------------
%% Titulos preliminares (Resumen, Contenido, Listas, Introduccion, Bibliografia)
%%-----------------------------------------------------------------------------
% Titulo de 20 pt: Introduccion, Lista de figuras, Lista de tablas,
% Lista de simbolos, Bibliografia.
\newcommand{\tituloPreliminar}[1]{%
  \cleardoublepage
  \phantomsection
  \markboth{#1}{#1}%
  \vspace*{100pt}%
  {\interlineadoSencillo\raggedright\bfseries\fontsize{20}{20}\selectfont #1\par}%
  \addcontentsline{toc}{chapter}{#1}%
  \vspace{10pt}%
  \par\nobreak\@afterheading
}

% Titulo de 18 pt: Resumen / Abstract (estilo "Titulos preliminares")
\newcommand{\tituloResumen}[1]{%
  \phantomsection
  \markboth{#1}{#1}%
  \vspace*{100pt}%
  {\interlineadoSencillo\raggedright\bfseries\fontsize{18}{18}\selectfont #1\par}%
  \addcontentsline{toc}{chapter}{#1}%
  \vspace{10pt}%
  \par\nobreak\@afterheading
}

% Version sin entrada en la tabla de contenido
\newcommand{\tituloPreliminarSinTDC}[1]{%
  \phantomsection
  \markboth{#1}{#1}%
  \vspace*{100pt}%
  {\interlineadoSencillo\raggedright\bfseries\fontsize{20}{20}\selectfont #1\par}%
  \vspace{10pt}%
  \par\nobreak\@afterheading
}

%%-----------------------------------------------------------------------------
%% Anexos: estilo "Subttulo / Titulo Anexos" -> Arial 21 pt negrilla, "A.", "B."
%%-----------------------------------------------------------------------------
\newcommand{\iniciarAnexos}{%
  \appendix
  \renewcommand{\thechapter}{\Alph{chapter}}%
  \titleformat{\chapter}[hang]
    {\normalfont\interlineadoSencillo\bfseries\fontsize{21}{21}\selectfont\raggedright}
    {\thechapter.}{0.5em}{}
  \titlespacing*{\chapter}{0pt}{\espAnteCapitulo}{\espTrasCapitulo}%
}

%%-----------------------------------------------------------------------------
%% Encabezados y pies
%%   Word: cada seccion tiene encabezado par e impar distintos, con borde
%%   inferior de 0,5 pt (sz=4), en Arial 11 pt sin negrilla y color automatico.
%%   Pagina impar: titulo del capitulo a la izquierda, numero a la derecha.
%%   Pagina par:   numero a la izquierda, titulo de la tesis a continuacion.
%%   Los pies de pagina del original estan vacios.
%%-----------------------------------------------------------------------------
\RequirePackage{fancyhdr}

\newcommand{\tituloTesis}{Titulo de la tesis o trabajo de investigacion}
\newcommand{\settituloTesis}[1]{\renewcommand{\tituloTesis}{#1}}

\newcommand{\fuenteEncabezado}{\normalfont\interlineadoSencillo\mdseries\fontsize{11}{13}\selectfont}

\fancypagestyle{tesis}{%
  \fancyhf{}%
  \fancyhead[LE]{\fuenteEncabezado\thepage\hspace{1em}\nouppercase{\tituloTesis}}%
  \fancyhead[LO]{\fuenteEncabezado\nouppercase{\leftmark}}%
  \fancyhead[RO]{\fuenteEncabezado\thepage}%
  \renewcommand{\headrulewidth}{0.5pt}%
  \renewcommand{\footrulewidth}{0pt}%
}

% Primera pagina de cada capitulo/seccion: Word usa <w:titlePg/> (sin encabezado)
\fancypagestyle{plain}{%
  \fancyhf{}%
  \renewcommand{\headrulewidth}{0pt}%
  \renewcommand{\footrulewidth}{0pt}%
}

\pagestyle{tesis}

% Las paginas de relleno generadas por \cleardoublepage quedan totalmente
% en blanco (en el original de Word no llevan encabezado ni numero).
\renewcommand{\cleardoublepage}{%
  \clearpage
  \if@twoside
    \ifodd\c@page\else
      \hbox{}\thispagestyle{empty}\newpage
      \if@twocolumn\hbox{}\newpage\fi
    \fi
  \fi
}

% Marca de capitulo sin mayusculas forzadas y sin la palabra "Capitulo"
\renewcommand{\chaptermark}[1]{\markboth{\ifnum\value{secnumdepth}>-1 \thechapter.\ \fi #1}{}}
\renewcommand{\sectionmark}[1]{}

%%-----------------------------------------------------------------------------
%% Numeracion de paginas
%%   Word: preliminares en numeros romanos MAYUSCULAS; desde la Introduccion,
%%   arabigos reiniciando en 1. Cubierta y portada no se numeran pero se cuentan.
%%-----------------------------------------------------------------------------
\renewcommand{\frontmatter}{%
  \cleardoublepage
  \@mainmatterfalse
  \pagenumbering{Roman}%
}

%%-----------------------------------------------------------------------------
%% Figuras y tablas
%%   Numeracion por capitulo con guion de no separacion: "Figura 2-1:"
%%   Rotulo en negrilla, ENCIMA del objeto, texto justificado.
%%-----------------------------------------------------------------------------
\RequirePackage{graphicx}
\RequirePackage{float}
\RequirePackage[labelsep=space]{caption}

% Guion de no separacion (Word usa U+2011 en "Figura 2-1")
\DeclareRobustCommand{\guionNB}{\mbox{-}}
\renewcommand{\thefigure}{\thechapter\guionNB\arabic{figure}}
\renewcommand{\thetable}{\thechapter\guionNB\arabic{table}}

\DeclareCaptionLabelFormat{rotuloFigura}{\textbf{#1 #2:}}
\DeclareCaptionLabelFormat{rotuloTabla}{\textbf{#1 #2}:}

\captionsetup[figure]{%
  labelformat=rotuloFigura,
  labelsep=space,
  justification=justified,
  singlelinecheck=false,
  position=above,
  skip=6pt,
  font={normalsize,stretch=1.0}
}
\captionsetup[table]{%
  labelformat=rotuloTabla,
  labelsep=space,
  justification=justified,
  singlelinecheck=false,
  position=above,
  skip=6pt,
  font={normalsize,stretch=1.0}
}

\RequirePackage{array}
\RequirePackage{longtable}
\RequirePackage{tabularx}
% Word: "los titulos y datos se deben ubicar entre lineas horizontales y
% verticales cerradas" -> tablas con cuadricula completa (Table Grid, 0,5 pt).
\setlength{\arrayrulewidth}{0.5pt}

%%-----------------------------------------------------------------------------
%% Ecuaciones: numeradas por capitulo -> (2.1)
%%-----------------------------------------------------------------------------
\RequirePackage{amsmath,amssymb}
\numberwithin{equation}{chapter}

%%-----------------------------------------------------------------------------
%% Vinetas cuadradas, sin sangria y justificadas
%%-----------------------------------------------------------------------------
\RequirePackage{enumitem}
\setlist[itemize]{%
  label=\raisebox{0.15ex}{\scalebox{0.55}{$\blacksquare$}},
  leftmargin=1.2em,
  labelsep=0.6em,
  topsep=\baselineskip,
  itemsep=\baselineskip,
  parsep=0pt,
  partopsep=0pt
}
\RequirePackage{graphics}

%%-----------------------------------------------------------------------------
%% Notas al pie
%%   "interlineado sencillo, texto justificado y usando la letra del texto
%%    en 10 puntos"
%%-----------------------------------------------------------------------------
\RequirePackage[hang,flushmargin]{footmisc}
\renewcommand{\footnotelayout}{\interlineadoSencillo\fontsize{10}{12}\selectfont\justifying}
\setlength{\footnotesep}{0.6\baselineskip}

%%-----------------------------------------------------------------------------
%% Tabla de contenido, lista de figuras y lista de tablas
%%   TDC1: negrilla, con lineas de puntos.
%%   TDC2/TDC3: normal, con sangria.
%%-----------------------------------------------------------------------------
\RequirePackage{tocloft}

\renewcommand{\cfttoctitlefont}{\bfseries\fontsize{20}{20}\selectfont}
\renewcommand{\cftloftitlefont}{\bfseries\fontsize{20}{20}\selectfont}
\renewcommand{\cftlottitlefont}{\bfseries\fontsize{20}{20}\selectfont}
\renewcommand{\cftaftertoctitle}{\hfill}
\renewcommand{\cftafterloftitle}{\hfill}
\renewcommand{\cftafterlottitle}{\hfill}
\setlength{\cftbeforetoctitleskip}{100pt}
\setlength{\cftbeforeloftitleskip}{100pt}
\setlength{\cftbeforelottitleskip}{100pt}
\setlength{\cftaftertoctitleskip}{10pt}
\setlength{\cftafterloftitleskip}{10pt}
\setlength{\cftafterlottitleskip}{10pt}

\renewcommand{\cftdotsep}{1.5}
\renewcommand{\cftchapaftersnum}{.}
\setlength{\cftchapnumwidth}{1.6em}
\renewcommand{\cftchapfont}{\bfseries}
\renewcommand{\cftchappagefont}{\bfseries}
\renewcommand{\cftchapleader}{\bfseries\cftdotfill{\cftdotsep}}
\setlength{\cftbeforechapskip}{10pt}
\setlength{\cftbeforesecskip}{0pt}
\setlength{\cftbeforesubsecskip}{0pt}
\renewcommand{\cftsecfont}{\normalfont}
\renewcommand{\cftsubsecfont}{\normalfont}
\setlength{\cftsecindent}{0.4em}
\setlength{\cftsubsecindent}{1.1em}
\setlength{\cftfigindent}{0pt}
\setlength{\cfttabindent}{0pt}
\renewcommand{\cftfigfont}{\normalfont}
\renewcommand{\cfttabfont}{\normalfont}
% En el original las entradas se leen "Figura 2-1: Tipos y parte..." y
% "Tabla 2-1: Participacion...", con el rotulo en negrilla.
% Se escriben literalmente para que babel no pueda sobrescribirlos.
\renewcommand{\cftfigpresnum}{\bfseries Figura~}
\renewcommand{\cftfigaftersnum}{\bfseries:\space}
\setlength{\cftfignumwidth}{5.8em}
\renewcommand{\cfttabpresnum}{\bfseries Tabla~}
\renewcommand{\cfttabaftersnum}{\bfseries:\space}
\setlength{\cfttabnumwidth}{5.2em}

% Nombres en espanol tal como aparecen en la plantilla
\AtBeginDocument{%
  \renewcommand{\contentsname}{Contenido}%
  \renewcommand{\listfigurename}{Lista de figuras}%
  \renewcommand{\listtablename}{Lista de tablas}%
  \renewcommand{\bibname}{Bibliograf\'\i a}%
  \renewcommand{\figurename}{Figura}%
  \renewcommand{\tablename}{Tabla}%
  \renewcommand{\appendixname}{Anexo}%
  \renewcommand{\chaptername}{Cap\'\i tulo}%
}
% Babel puede reponer sus propios rotulos al seleccionar idioma; se vuelven a
% fijar al comenzar el cuerpo del documento.
\AtBeginDocument{%
  \renewcommand{\figurename}{Figura}%
  \renewcommand{\tablename}{Tabla}%
}

% Encabezado "Pag." sobre la primera entrada de cada lista
\newcommand{\encabezadoPag}{\hfill\normalfont\textbf{P\'ag.}\par\vspace{6pt}}

%%-----------------------------------------------------------------------------
%% Portada / cubierta
%%   Estilo "Ttulo" (Title): Arial 24 pt negrilla, centrado.
%%-----------------------------------------------------------------------------
\newcommand{\tituloPortada}[1]{%
  {\interlineadoSencillo\centering\bfseries\fontsize{24}{28}\selectfont #1\par}%
}
% Estilo "Textoennegrita" (alias "Texto de la tesis"): Arial 12 pt
\newcommand{\textoPortada}[1]{{\bfseries\fontsize{12}{14}\selectfont #1}}

\newenvironment{paginaPreliminar}%
  {\thispagestyle{empty}\interlineadoSencillo\centering}%
  {\par\clearpage}

%%-----------------------------------------------------------------------------
%% Bibliografia
%%   "Cada referencia bibliografica se inicia contra el margen izquierdo".
%%   Se elimina el titulo automatico: el titulo lo pone \tituloPreliminar.
%%-----------------------------------------------------------------------------
\renewenvironment{thebibliography}[1]{%
  \interlineadoSencillo
  \list{\@biblabel{\@arabic\c@enumiv}}%
       {\setlength{\labelwidth}{0pt}%
        \setlength{\labelsep}{0.5em}%
        \setlength{\leftmargin}{0pt}%
        \setlength{\itemindent}{0pt}%
        \setlength{\itemsep}{0.5\baselineskip}%
        \setlength{\parsep}{0pt}%
        \usecounter{enumiv}%
        \let\p@enumiv\@empty
        \renewcommand\theenumiv{\@arabic\c@enumiv}}%
  \sloppy
  \clubpenalty4000 \@clubpenalty\clubpenalty \widowpenalty4000%
  \sfcode`\.\@m
}{%
  \def\@noitemerr{\@latex@warning{Empty `thebibliography' environment}}%
  \endlist
}

%%-----------------------------------------------------------------------------
%% Hiperenlaces (sin color, para que la impresion sea identica a Word)
%%-----------------------------------------------------------------------------
\RequirePackage[hidelinks,breaklinks=true]{hyperref}
\RequirePackage{url}

%%-----------------------------------------------------------------------------
%% Evitar titulos o renglones sueltos al final de pagina
%%   "Se debe evitar titulos o subtitulos solos al final de la pagina o
%%    renglones sueltos" -> equivalente a keepNext / widowControl de Word.
%%-----------------------------------------------------------------------------
\clubpenalty=10000
\widowpenalty=10000
\displaywidowpenalty=10000
\raggedbottom

\endinput
