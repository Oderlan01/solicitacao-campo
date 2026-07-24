(vl-load-com)

;;; ================================================================
;;; stw-automacao.lsp
;;; Ferramentas AutoCAD para sincronizacao de atributos com Excel
;;; e propagacao de tags entre blocos pai/filho.
;;;
;;; FLUXO UNIFICADO (arquivo unico Desktop\todos_atributos.csv):
;;;
;;;   CAD -> Excel:
;;;     1. AutoCAD: STWExportar    -> Desktop\todos_atributos.csv
;;;     2. Excel:   ImportarDoCAD  <- Desktop\todos_atributos.csv
;;;
;;;   Excel -> CAD:
;;;     1. Excel:   ExportarParaCAD -> Desktop\todos_atributos.csv
;;;     2. AutoCAD: STWImportar     <- Desktop\todos_atributos.csv
;;;
;;; Comandos disponiveis (todos comecam com STW):
;;;   STWSelecionarPai    -> usuario clica no bloco pai; so os filhos dele sao atualizados
;;;   STWAtualizarTags    -> atualiza todos os blocos do desenho (pai/filho + ID_VISIVEL)
;;;   STWExportar         -> atualiza tudo e exporta CSV (CAD -> Excel)
;;;   STWImportar         <- importa CSV e atualiza blocos (Excel -> CAD)
;;;   STWExportarDominios -> exporta dominios.csv (listas de validacao p/ dropdowns Excel)
;;;
;;; Compativel com Modulo3_SincronizarCAD.bas (VBA Excel):
;;;   ExportarParaCAD() atualiza todos_atributos.csv em memoria
;;;   ImportarDoCAD()   le todos_atributos.csv gerado pelo STWExportar
;;;
;;; ARQUITETURA (hibrida, centrada no LISP):
;;;   O LISP e dono dos DOMINIOS de validacao (STW:DOMINIOS), exportados para
;;;   o Excel montar os dropdowns sem listas hard-coded. O dicionario canonico
;;;   de variaveis (atributo CAD <-> nome amigavel <-> grupo <-> direcao) vive
;;;   em FLUXO_DADOS.md, que tambem documenta as regras de negocio.
;;; ================================================================


;; ================================================================
;; DOMINIOS DE VALIDACAO  (listas de dropdown dependentes)
;;
;; Cada entrada: (CAMPO  DEPENDE_DE  REGRAS)
;;   CAMPO      : campo cujo dropdown estamos definindo (nome amigavel)
;;   DEPENDE_DE : campo do qual o dropdown depende (ou nil)
;;   REGRAS     : lista de (VALORES_PAI . OPCOES)
;;                VALORES_PAI : lista de valores do campo pai que disparam
;;                              estas OPCOES; "*" = qualquer outro valor
;;                OPCOES      : lista de strings oferecidas no dropdown
;; Exportado por STWExportarDominios para dominios.csv.
;;
;; Nota de codificacao: valores mantidos SEM acento (ASCII) para evitar
;; mojibake entre AutoCAD (ANSI/1252) e a leitura ANSI do VBA. O dropdown de
;; INTERFACE depende do valor escolhido no dropdown de ACIONAMENTO (ambos
;; vindos daqui), entao a dependencia permanece consistente.
;; ================================================================
(setq STW:DOMINIOS
  '(("ACIONAMENTO" "CATEGORIA"
      (("MOTOR") .
        ("Soft Starter" "Inversor" "Partida Direta" "Partida Direta + Inv"
         "Inversor + PD" "Partida Direta + Rev" "Partida Inteligente"
         "Soft Starter + Rev"))
      (("*") .
        ("Capacitivo" "Indutivo" "Magnetico" "Trava Seg." "Laser"
         "Temperatura" "Ultrassonico" "Contato" "Acionamento" "Feedback"
         "Valvula" "Simples Solenoide" "Dupla Solenoide")))
    ("INTERFACE" "ACIONAMENTO"
      (("Inversor" "Soft Starter" "Partida Inteligente" "Soft Starter + Rev"
        "Inversor + PD" "Partida Direta + Inv") .
        ("I/O" "Rede"))
      (("Partida Direta" "Partida Direta + Rev") .
        ("I/O"))
      (("Capacitivo" "Indutivo" "Magnetico" "Trava Seg." "Laser"
        "Ultrassonico" "Contato" "Feedback") .
        ("DI" "AI" "Circuito Eletrico" "I/O Link"))
      (("Temperatura") .
        ("AI" "I/O Link" "RTD"))
      (("Acionamento" "Valvula" "Dupla Solenoide" "Simples Solenoide") .
        ("AO" "DO" "Circuito Eletrico" "I/O Link")))))


;; ================================================================
;; HELPERS GLOBAIS
;; ================================================================

;; Remove \r final (CRLF Windows/Excel)
(defun STW:strip-cr (s)
  (if (and (> (strlen s) 0)
           (= (substr s (strlen s) 1) "\r"))
    (substr s 1 (1- (strlen s)))
    s))

;; Separa string por delimitador com tratamento de CRLF
(defun STW:quebrar (str delim / pos lst)
  (setq str (STW:strip-cr str))
  (while (setq pos (vl-string-search delim str))
    (setq lst (cons (substr str 1 pos) lst))
    (setq str (substr str (+ pos 2))))
  (reverse (cons str lst)))

;; Retorna o nome efetivo de um VLA INSERT (suporta dinamicos).
;; Garante retorno de string ou nil — nunca T ou outro tipo.
(defun STW:nome-efetivo (obj / nome)
  (setq nome nil)
  (if (vlax-property-available-p obj 'EffectiveName)
    (setq nome (vlax-get-property obj 'EffectiveName)))
  (if (not (and nome (= (type nome) 'STR) (/= nome "")))
    (setq nome (vlax-get-property obj 'Name)))
  (if (= (type nome) 'STR) nome nil))

;; Retorna o caminho da Area de Trabalho (Desktop)
(defun STW:desktop ()
  (vl-registry-read
    "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders"
    "Desktop"))


;; ================================================================
;; NUCLEO: STW:atualizar-todos
;; Percorre todos os blocos ATL_STW e executa:
;;   1. Propaga 0E_TAG do pai para os filhos aninhados (base ATTDEF + pai)
;;   2. Sincroniza ID_VISIVEL = handle do bloco
;;   3. Sincroniza NOME_DO_BLOCO = nome efetivo do bloco
;; Retorna lista (total-tag-pai total-id-sync) para relatorio.
;; ================================================================
(defun STW:atualizar-todos (/ ss i obj nomeReal hnd listaAtribs atrib
                               nomeAtrib tagValor n
                               totalTagPai totalIdSync)
  (setq ss          (ssget "X" '((0 . "INSERT")))
        totalTagPai 0
        totalIdSync 0)
  (if (not ss)
    (progn (princ "\nNenhum bloco encontrado.") (list 0 0))
    (progn
      ;; --- Passo 1: propaga 0E_TAG pai -> filhos ---
      (setq i 0)
      (while (< i (sslength ss))
        (setq obj      (vlax-ename->vla-object (ssname ss i))
              nomeReal (STW:nome-efetivo obj))
        (if (and nomeReal (= (type nomeReal) 'STR)
                 (>= (strlen nomeReal) 7)
                 (= (strcase (substr nomeReal 1 7)) "ATL_STW"))
          (progn
            (setq tagValor (STW:get-0e-tag obj))
            (if (and tagValor (= (type tagValor) 'STR) (/= tagValor "") (/= tagValor "-"))
              (progn
                (setq n (STW:modificar-def nomeReal tagValor))
                (setq totalTagPai (+ totalTagPai n))))))
        (setq i (1+ i)))

      ;; --- Passo 2: sincroniza ID_VISIVEL e NOME_DO_BLOCO ---
      (setq i 0)
      (while (< i (sslength ss))
        (setq obj      (vlax-ename->vla-object (ssname ss i))
              nomeReal (STW:nome-efetivo obj)
              hnd      (vla-get-Handle obj))
        (if (and nomeReal (= (type nomeReal) 'STR)
                 (= (strcase (substr nomeReal 1 7)) "ATL_STW")
                 (= (vla-get-HasAttributes obj) :vlax-true))
          (progn
            (setq listaAtribs
                  (vlax-safearray->list (vlax-variant-value (vla-GetAttributes obj))))
            (foreach atrib listaAtribs
              (setq nomeAtrib (strcase (vla-get-TagString atrib)))
              (cond
                ((= nomeAtrib "ID_VISIVEL")
                 (vla-put-TextString atrib hnd)
                 (setq totalIdSync (1+ totalIdSync)))
                ((= nomeAtrib "NOME_DO_BLOCO")
                 (vla-put-TextString atrib nomeReal))))
            (vla-update obj)))
        (setq i (1+ i)))

      (list totalTagPai totalIdSync))))


;; ================================================================
;; COMANDO 1: STWAtualizarTags
;; Atualiza todos os blocos ATL_STW no desenho:
;;   - Propaga 0E_TAG do pai para os filhos aninhados
;;   - Sincroniza ID_VISIVEL com o handle real do bloco
;;   - Sincroniza NOME_DO_BLOCO com o nome efetivo
;; STWExportar chama este comando automaticamente antes de gravar.
;; ================================================================
(defun c:STWAtualizarTags (/ resultado)
  (princ "\nAtualizando tags...")
  (setq resultado (STW:atualizar-todos))
  (command "_.REGENALL")
  (princ (strcat "\nSTWAtualizarTags concluido:"
                 "\n  0E_TAG filhos atualizados : " (itoa (car resultado))
                 "\n  ID_VISIVEL sincronizados  : " (itoa (cadr resultado))))
  (princ))


;; ================================================================
;; HELPER: STW:atribs-visiveis
;;
;; Coleta atributos visiveis de um INSERT (modelo ou definicao de bloco).
;; Usa o flag DXF 70 (bit 0 = invisivel) — mais confiavel que VLA Invisible.
;; Exclui automaticamente ID_VISIVEL e NOME_DO_BLOCO (atributos internos).
;;
;; Retorna:  cons (novas-tags . lista-de-pares-tag-valor)
;;   novas-tags : tags ainda nao presentes em listaTags (para agregar)
;;   lista-pares: lista de (tag . valor) dos atributos coletados
;; ================================================================
(defun STW:atribs-visiveis (insEnt listaTags / subEnt sd flags tag val result novas)
  (setq subEnt (entnext insEnt)
        result nil
        novas  nil)
  (while (and subEnt
              (setq sd (entget subEnt))
              (= (cdr (assoc 0 sd)) "ATTRIB"))
    (setq flags (cdr (assoc 70 sd))
          tag   (cdr (assoc 2  sd))
          val   (cdr (assoc 1  sd)))
    ;; bit 0 do flag 70 = invisivel; 0 = visivel
    (if (and (zerop (logand (if flags flags 0) 1))
             (not (member (strcase tag) '("ID_VISIVEL" "NOME_DO_BLOCO"))))
      (progn
        (if (and (not (member tag listaTags))
                 (not (member tag novas)))
          (setq novas (cons tag novas)))
        (setq result (cons (cons tag val) result))))
    (setq subEnt (entnext subEnt)))
  (cons novas result))


;; ================================================================
;; COMANDO 2: STWExportar
;;
;; Atualiza todos os atributos e exporta para CSV.
;; Arquivo: Desktop\todos_atributos.csv
;; Formato: HANDLE_CAD;Nome_Bloco;TAG1;TAG2;...  (tags em ordem alfabetica)
;;
;; Correcoes aplicadas:
;;   [1] Visibilidade via flag DXF 70 (bit 0), nao VLA Invisible
;;   [2] ID_VISIVEL e NOME_DO_BLOCO excluidos das colunas de atributo
;;   [3] Blocos ATL_STW aninhados na definicao do bloco pai tambem
;;       sao exportados como linhas proprias (com seu proprio HANDLE)
;; ================================================================
(defun c:STWExportar (/ caminho ss i ent obj nomeReal hnd
                        blocosDados listaTags cabecalho linhaTexto
                        busca ret parAtribs filhoAttribs
                        blkRec defEnt defData nomeFilho hndFilho arq)
  (princ "\nAtualizando tags antes de exportar...")
  (STW:atualizar-todos)

  (setq caminho     (strcat (STW:desktop) "\\todos_atributos.csv")
        ss          (ssget "X" '((0 . "INSERT")))
        blocosDados nil
        listaTags   nil
        i           0)

  (if (not ss)
    (princ "\nNenhum bloco encontrado.")
    (progn
      (while (< i (sslength ss))
        (setq ent      (ssname ss i)
              obj      (vlax-ename->vla-object ent)
              nomeReal (STW:nome-efetivo obj)
              hnd      (vla-get-Handle obj))

        (if (and nomeReal
                 (= (type nomeReal) 'STR)
                 (>= (strlen nomeReal) 7)
                 (= (strcase (substr nomeReal 1 7)) "ATL_STW"))
          (progn
            ;; --- [1][2] Atributos do bloco PAI (visiveis, sem internos) ---
            (setq ret       (STW:atribs-visiveis ent listaTags)
                  listaTags (append listaTags (car ret))
                  parAtribs (cdr ret))
            (setq blocosDados (cons (list hnd nomeReal parAtribs) blocosDados))

            ;; --- [3] Blocos FILHO aninhados na definicao do bloco pai ---
            (setq blkRec (tblsearch "BLOCK" nomeReal))
            (if blkRec
              (progn
                (setq defEnt (cdr (assoc -2 blkRec)))
                (while defEnt
                  (setq defData (entget defEnt))
                  (if (= (cdr (assoc 0 defData)) "INSERT")
                    (progn
                      (setq nomeFilho (cdr (assoc 2 defData))
                            hndFilho  (cdr (assoc 5 defData)))
                      (if (and nomeFilho hndFilho
                               (>= (strlen nomeFilho) 7)
                               (= (strcase (substr nomeFilho 1 7)) "ATL_STW"))
                        (progn
                          (setq ret         (STW:atribs-visiveis defEnt listaTags)
                                listaTags   (append listaTags (car ret))
                                filhoAttribs (cdr ret))
                          (setq blocosDados
                                (cons (list hndFilho nomeFilho filhoAttribs)
                                      blocosDados))))))
                  (setq defEnt (entnext defEnt)))))))
        (setq i (1+ i)))

      ;; --- Escrever CSV ---
      (setq listaTags (acad_strlsort listaTags)
            arq       (open caminho "w")
            cabecalho "HANDLE_CAD;Nome_Bloco")
      (foreach tag listaTags
        (setq cabecalho (strcat cabecalho ";" tag)))
      (write-line cabecalho arq)
      (foreach blk blocosDados
        (setq hnd        (car   blk)
              nomeReal   (cadr  blk)
              parAtribs  (caddr blk)
              linhaTexto (strcat hnd ";" nomeReal))
        (foreach tag listaTags
          (setq busca (assoc tag parAtribs))
          (setq linhaTexto (strcat linhaTexto ";" (if busca (cdr busca) ""))))
        (write-line linhaTexto arq))
      (close arq)
      (command "_.REGEN")
      (princ (strcat "\nSTWExportar concluido: " caminho))))
  (princ))


;; ================================================================
;; COMANDO 3: STWImportar
;; Importa Desktop\todos_atributos.csv e atualiza os blocos.
;; Formato: HANDLE_CAD;Nome_Bloco;TAG1;TAG2;...
;; Nao sobrescreve atributos com valor vazio no CSV.
;; ================================================================
(defun c:STWImportar (/ caminho arq linha cabecalho listaTags dados
                        hndVal ent obj listaAtribs i tagAtual
                        valorNovo atrib atualizados naoEncontrados)
  (setq caminho (strcat (STW:desktop) "\\todos_atributos.csv"))
  (setq arq (open caminho "r"))
  (if (not arq)
    (princ (strcat "\nArquivo nao encontrado: " caminho
                   "\nExecute primeiro ExportarParaCAD no Excel."))
    (progn
      (setq linha          (read-line arq)
            cabecalho      (STW:quebrar linha ";")
            listaTags      (cddr cabecalho) ; pula HANDLE_CAD e Nome_Bloco
            atualizados    0
            naoEncontrados 0)
      (while (setq linha (read-line arq))
        (setq dados  (STW:quebrar linha ";")
              hndVal (car dados)
              ent    (handent hndVal))
        (if (not ent)
          (setq naoEncontrados (1+ naoEncontrados))
          (progn
            (setq obj (vlax-ename->vla-object ent))
            (if (= (vla-get-HasAttributes obj) :vlax-true)
              (progn
                (setq listaAtribs
                      (vlax-safearray->list (vlax-variant-value (vla-GetAttributes obj))))
                (setq i 0)
                (while (< i (length listaTags))
                  (setq tagAtual  (nth i listaTags)
                        valorNovo (nth (+ i 2) dados))
                  (if (and valorNovo (/= valorNovo ""))
                    (foreach atrib listaAtribs
                      (if (= (strcase (vla-get-TagString atrib)) (strcase tagAtual))
                        (progn
                          (vla-put-TextString atrib valorNovo)
                          (setq atualizados (1+ atualizados))))))
                  (setq i (1+ i)))
                (vla-update obj))))))
      (close arq)
      (command "_.REGEN")
      (princ (strcat "\nSTWImportar concluido: "
                     (itoa atualizados) " atributo(s) atualizado(s)"
                     (if (> naoEncontrados 0)
                       (strcat ", " (itoa naoEncontrados) " handle(s) nao encontrado(s).")
                       ".")))))
  (princ))


;; ================================================================
;; HELPERS: propagacao 0E_TAG pai -> filhos
;;
;; Estrategia anti-acumulacao com XDATA:
;;   - Na primeira execucao: usa o valor atual do ATTRIB como prefixo
;;     (preserva VP1-, VP2-, VP3- definidos manualmente) e salva em XDATA
;;   - Nas execucoes seguintes: le o prefixo do XDATA (fixo, nunca muda)
;;     e aplica o novo tag do pai por cima
;;
;; Isso garante que trocar "SL-agora" -> "SL-novo" no pai resulta em
;; "VP1-SL-novo" e nao "VP1-SL-agoraSL-novo".
;; ================================================================

;; Le o prefixo armazenado em XDATA no ATTRIB.
;; Retorna string ou nil.
(defun STW:ler-prefixo (subEnt / dados xd appd val)
  (setq dados (entget subEnt '("STW_TAG"))
        val   nil)
  (if dados
    (progn
      (setq xd (assoc -3 dados))
      (if xd
        (progn
          (setq appd (assoc "STW_TAG" (cdr xd)))
          (if appd
            (progn
              (setq val (cdr (assoc 1000 (cdr appd))))
              (if (not (and val (= (type val) 'STR)))
                (setq val nil))))))))
  val)

;; Salva o prefixo em XDATA no ATTRIB para uso nas proximas execucoes
(defun STW:salvar-prefixo (subDados prefixo)
  (regapp "STW_TAG")
  (setq subDados (vl-remove-if (function (lambda (x) (= (car x) -3))) subDados))
  (append subDados (list (list -3 (list "STW_TAG" (cons 1000 prefixo))))))

;; Retorna o valor do atributo 0E_TAG de um VLA INSERT
(defun STW:get-0e-tag (obj / val)
  (setq val nil)
  (if (= (vla-get-HasAttributes obj) :vlax-true)
    (foreach att (vlax-safearray->list
                   (vlax-variant-value (vla-GetAttributes obj)))
      (if (= (strcase (vla-get-TagString att)) "0E_TAG")
        (setq val (vla-get-TextString att)))))
  val)

;; Modifica na DEFINICAO do bloco pai os atributos dos filhos aninhados:
;;   0E_TAG       = prefixo (XDATA ou valor atual) + 0E_TAG do pai
;;   ID_VISIVEL   = handle do INSERT filho dentro da definicao
;;   NOME_DO_BLOCO = nome do bloco filho
;; Blocos standalone (sem filhos aninhados) nao sao afetados.
;; Retorna quantidade de ATTRIBs 0E_TAG alterados.
(defun STW:modificar-def (blkname valorPai / blkRec blkEnt ent dados tipo
                           nomeFilho hndFilho prefixo novoValor
                           subEnt subDados tag modificou n)
  (setq blkRec (tblsearch "BLOCK" blkname) n 0)
  (if blkRec
    (progn
      (setq blkEnt (cdr (assoc -1 blkRec))
            ent    (cdr (assoc -2 blkRec)))
      (while ent
        (setq dados (entget ent)
              tipo  (cdr (assoc 0 dados)))
        (if (= tipo "INSERT")
          (progn
            (setq nomeFilho (cdr (assoc 2 dados))
                  hndFilho  (cdr (assoc 5 dados))
                  subEnt    (entnext ent)
                  modificou nil)
            (while (and subEnt
                        (setq subDados (entget subEnt))
                        (= (cdr (assoc 0 subDados)) "ATTRIB"))
              (setq tag (strcase (cdr (assoc 2 subDados))))
              (cond
                ((= tag "0E_TAG")
                 (setq prefixo (STW:ler-prefixo subEnt))
                 (if (not prefixo)
                   (progn
                     (setq prefixo (cdr (assoc 1 subDados)))
                     (if (not (and prefixo (= (type prefixo) 'STR)))
                       (setq prefixo ""))))
                 (setq novoValor (strcat prefixo valorPai))
                 (setq subDados (STW:salvar-prefixo
                                  (subst (cons 1 novoValor) (assoc 1 subDados) subDados)
                                  prefixo))
                 (entmod subDados)
                 (setq modificou T n (1+ n)))
                ((= tag "ID_VISIVEL")
                 (if hndFilho
                   (progn
                     (entmod (subst (cons 1 hndFilho) (assoc 1 subDados) subDados))
                     (setq modificou T))))
                ((= tag "NOME_DO_BLOCO")
                 (entmod (subst (cons 1 nomeFilho) (assoc 1 subDados) subDados))
                 (setq modificou T)))
              (setq subEnt (entnext subEnt)))
            (if modificou (entupd ent))))
        (setq ent (entnext ent)))
      (if (and blkEnt (> n 0)) (entupd blkEnt))))
  n)


;; ================================================================
;; COMANDO 4: STWSelecionarPai
;; O usuario clica em um bloco pai no desenho.
;; Apenas os filhos daquele bloco sao atualizados:
;;   - 0E_TAG filho = prefixo salvo + 0E_TAG do pai selecionado
;;   - ID_VISIVEL e NOME_DO_BLOCO dos filhos tambem atualizados
;; Todos os outros blocos do desenho ficam intocados.
;;
;; Uso tipico: bloco pai recem criado ou com 0E_TAG alterado no CAD.
;; ================================================================
(defun c:STWSelecionarPai (/ sel ent obj nomeReal tagValor n hnd listaAtribs atrib nomeAtrib)
  (princ "\nClique no bloco pai para atualizar seus filhos: ")
  (setq sel (entsel))
  (if (not sel)
    (princ "\nNenhuma entidade selecionada.")
    (progn
      (setq ent      (car sel)
            obj      (vlax-ename->vla-object ent)
            nomeReal (STW:nome-efetivo obj))
      (cond
        ((not (and nomeReal (= (type nomeReal) 'STR)))
         (princ "\nEntidade invalida. Selecione um bloco ATL_STW."))
        ((not (= (strcase (substr nomeReal 1 7)) "ATL_STW"))
         (princ (strcat "\nBloco '" nomeReal "' nao e um bloco ATL_STW.")))
        ((not (= (vla-get-HasAttributes obj) :vlax-true))
         (princ (strcat "\nBloco '" nomeReal "' nao possui atributos.")))
        (T
         (setq tagValor (STW:get-0e-tag obj))
         (if (not (and tagValor (= (type tagValor) 'STR) (/= tagValor "") (/= tagValor "-")))
           (princ (strcat "\nBloco '" nomeReal "' nao tem 0E_TAG definido."))
           (progn
             (setq n (STW:modificar-def nomeReal tagValor))
             (if (> n 0)
               (progn
                 (setq hnd        (vla-get-Handle obj)
                       listaAtribs (vlax-safearray->list
                                     (vlax-variant-value (vla-GetAttributes obj))))
                 (foreach atrib listaAtribs
                   (setq nomeAtrib (strcase (vla-get-TagString atrib)))
                   (cond
                     ((= nomeAtrib "ID_VISIVEL")
                      (vla-put-TextString atrib hnd))
                     ((= nomeAtrib "NOME_DO_BLOCO")
                      (vla-put-TextString atrib nomeReal))))
                 (vla-update obj)
                 (command "_.REGENALL")
                 (princ (strcat "\nSTWSelecionarPai concluido: '"
                                nomeReal "' | 0E_TAG='" tagValor
                                "' | " (itoa n) " filho(s) atualizado(s).")))
               (princ (strcat "\nBloco '" nomeReal
                              "' nao tem filhos com 0E_TAG para atualizar."))))))))
  (princ))


;; ================================================================
;; COMANDO 5: STWExportarDominios
;;
;; Exporta as listas de validacao (STW:DOMINIOS) para um CSV que o Excel
;; consome para montar os dropdowns dependentes (RN4), sem listas hard-coded.
;;
;; Arquivo: Desktop\dominios.csv
;; Formato: CAMPO;DEPENDE_DE;VALOR_PAI;OPCOES
;;   - uma linha por par (regra) do dominio
;;   - VALOR_PAI: valores do campo pai separados por "|" ("*" = demais casos)
;;   - OPCOES   : opcoes do dropdown separadas por "|"
;; ================================================================
(defun c:STWExportarDominios (/ caminho arq dom campo dependeDe regra
                                valoresPai opcoes linhaTexto nLinhas)
  (setq caminho (strcat (STW:desktop) "\\dominios.csv")
        arq     (open caminho "w")
        nLinhas 0)
  (if (not arq)
    (princ (strcat "\nNao foi possivel criar: " caminho))
    (progn
      (write-line "CAMPO;DEPENDE_DE;VALOR_PAI;OPCOES" arq)
      (foreach dom STW:DOMINIOS
        (setq campo     (car dom)
              dependeDe  (cadr dom))
        (foreach regra (cddr dom)
          (setq valoresPai (STW:juntar (car regra) "|")
                opcoes      (STW:juntar (cdr regra) "|")
                linhaTexto  (strcat campo ";"
                                    (if dependeDe dependeDe "") ";"
                                    valoresPai ";" opcoes))
          (write-line linhaTexto arq)
          (setq nLinhas (1+ nLinhas))))
      (close arq)
      (princ (strcat "\nSTWExportarDominios concluido: " caminho
                     " (" (itoa nLinhas) " regra(s))"))))
  (princ))

;; Junta uma lista de strings com um separador. "" se lista vazia.
(defun STW:juntar (lst sep / res)
  (setq res "")
  (foreach s lst
    (setq res (if (= res "") s (strcat res sep s))))
  res)


;; ================================================================
(princ "\nstw-automacao.lsp carregado. Comandos disponiveis:")
(princ "\n  STWAtualizarTags    -> atualiza todos os blocos (pai/filho + ID_VISIVEL)")
(princ "\n  STWExportar         -> atualiza tudo e exporta Desktop\\todos_atributos.csv")
(princ "\n  STWImportar         <- importa Desktop\\todos_atributos.csv (Excel -> CAD)")
(princ "\n  STWSelecionarPai    -> clique no bloco pai; so os filhos dele sao atualizados")
(princ "\n  STWExportarDominios -> exporta Desktop\\dominios.csv (dropdowns do Excel)")
(princ)
