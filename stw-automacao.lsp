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
;;;   STWAtualizarTags     -> atualiza 0E_TAG pai/filho + ID_VISIVEL + NOME_DO_BLOCO
;;;   STWExportar          -> atualiza tudo e exporta CSV (CAD -> Excel)
;;;   STWImportar          <- importa CSV e atualiza blocos (Excel -> CAD)
;;; ================================================================


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
                               nomeAtrib tagValor n blocosSync
                               totalTagPai totalIdSync)
  (setq ss          (ssget "X" '((0 . "INSERT")))
        totalTagPai 0
        totalIdSync 0
        blocosSync  '())
  (if (not ss)
    (progn (princ "\nNenhum bloco encontrado.") (list 0 0))
    (progn
      ;; --- Passo 1: propaga 0E_TAG pai -> filhos ---
      (setq i 0)
      (while (< i (sslength ss))
        (setq obj      (vlax-ename->vla-object (ssname ss i))
              nomeReal (STW:nome-efetivo obj))
        (if (and nomeReal (stringp nomeReal) (/= (substr nomeReal 1 1) "*"))
          (progn
            (setq tagValor (STW:get-0e-tag obj))
            (if (and tagValor (stringp tagValor) (/= tagValor "") (/= tagValor "-"))
              (progn
                (setq n (STW:modificar-def nomeReal tagValor))
                (setq totalTagPai (+ totalTagPai n))
                (if (and (> n 0) (not (member nomeReal blocosSync)))
                  (setq blocosSync (cons nomeReal blocosSync)))))))
        (setq i (1+ i)))
      (if blocosSync
        (foreach blk blocosSync (STW:sync-bloco blk)))

      ;; --- Passo 2: sincroniza ID_VISIVEL e NOME_DO_BLOCO ---
      (setq i 0)
      (while (< i (sslength ss))
        (setq obj      (vlax-ename->vla-object (ssname ss i))
              nomeReal (STW:nome-efetivo obj)
              hnd      (vla-get-Handle obj))
        (if (and nomeReal (stringp nomeReal)
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
;; COMANDO 2: STWExportar
;; Atualiza todos os atributos e exporta para CSV.
;; Arquivo: Desktop\todos_atributos.csv
;; Formato: HANDLE_CAD;Nome_Bloco;TAG1;TAG2;...
;; ================================================================
(defun c:STWExportar (/ caminho ss i ent obj nomeReal hnd
                        listaAtribs nomeAtrib valorAtrib arq
                        blocosDados listaTags cabecalho linhaTexto
                        busca parAtribs)
  ;; Garante que tudo esta atualizado antes de exportar
  (princ "\nAtualizando tags antes de exportar...")
  (STW:atualizar-todos)

  (setq caminho (strcat (STW:desktop) "\\todos_atributos.csv"))
  (setq ss (ssget "X" '((0 . "INSERT"))))
  (if (not ss)
    (princ "\nNenhum bloco encontrado.")
    (progn
      (setq blocosDados nil listaTags nil i 0)
      (while (< i (sslength ss))
        (setq ent      (ssname ss i)
              obj      (vlax-ename->vla-object ent)
              hnd      (vla-get-Handle obj)
              nomeReal (STW:nome-efetivo obj))
        (if (and nomeReal (stringp nomeReal)
                 (= (strcase (substr nomeReal 1 7)) "ATL_STW")
                 (= (vla-get-HasAttributes obj) :vlax-true))
          (progn
            (setq listaAtribs
                  (vlax-safearray->list (vlax-variant-value (vla-GetAttributes obj)))
                  parAtribs nil)
            (foreach atrib listaAtribs
              (if (= (vla-get-Invisible atrib) :vlax-false)
                (progn
                  (setq nomeAtrib  (vla-get-TagString atrib)
                        valorAtrib (vla-get-TextString atrib))
                  (if (not (member nomeAtrib listaTags))
                    (setq listaTags (cons nomeAtrib listaTags)))
                  (setq parAtribs (cons (cons nomeAtrib valorAtrib) parAtribs)))))
            (setq blocosDados (cons (list hnd nomeReal parAtribs) blocosDados))))
        (setq i (1+ i)))
      (setq listaTags (acad_strlsort listaTags))
      (setq arq (open caminho "w"))
      (setq cabecalho "HANDLE_CAD;Nome_Bloco")
      (foreach tag listaTags
        (setq cabecalho (strcat cabecalho ";" tag)))
      (write-line cabecalho arq)
      (foreach blk blocosDados
        (setq hnd       (car blk)
              nomeReal  (cadr blk)
              parAtribs (caddr blk)
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
;; COMANDO 2: STWImportar
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
      (setq linha         (read-line arq)
            cabecalho     (STW:quebrar linha ";")
            listaTags     (cddr cabecalho) ; pula HANDLE_CAD e Nome_Bloco
            atualizados   0
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
;; Retorna string ou nil — nunca T ou outro tipo nao-string.
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
  ; Remove -3 existente e adiciona novo com o prefixo
  (setq subDados (vl-remove-if '(lambda (x) (= (car x) -3)) subDados))
  (append subDados (list (list -3 (list "STW_TAG" (cons 1000 prefixo))))))

;; Retorna o valor do atributo 0E_TAG de um VLA INSERT no modelo
(defun STW:get-0e-tag (obj / val)
  (setq val nil)
  (if (= (vla-get-HasAttributes obj) :vlax-true)
    (foreach att (vlax-safearray->list
                   (vlax-variant-value (vla-GetAttributes obj)))
      (if (= (strcase (vla-get-TagString att)) "0E_TAG")
        (setq val (vla-get-TextString att)))))
  val)

;; Modifica na DEFINICAO do bloco pai os atributos dos filhos aninhados:
;;   0E_TAG      = prefixo (XDATA ou valor atual) + 0E_TAG do pai
;;   ID_VISIVEL  = handle do INSERT filho dentro da definicao
;;   NOME_DO_BLOCO = nome do bloco filho
;;
;; Blocos standalone (sem filhos aninhados) nao sao afetados.
;; Retorna quantidade de ATTRIBs 0E_TAG alterados.
(defun STW:modificar-def (blkname valorPai / blkRec blkEnt ent dados tipo
                           nomeFilho hndFilho prefixo novoValor
                           subEnt subDados tag modificou n)
  (setq blkRec (tblsearch "BLOCK" blkname) n 0)
  (if (not blkRec) (return n))
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
             ; Prefixo: le do XDATA (execucoes anteriores) ou do valor atual (1a vez)
             ; Garante que prefixo e sempre uma string (nunca T ou nil)
             (setq prefixo (STW:ler-prefixo subEnt))
             (if (not prefixo)
               (progn
                 (setq prefixo (cdr (assoc 1 subDados)))
                 (if (not (and prefixo (= (type prefixo) 'STR)))
                   (setq prefixo ""))))
             (setq novoValor (strcat prefixo valorPai))
             ; Salva prefixo em XDATA e atualiza valor em uma unica chamada entmod
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
  (if (and blkEnt (> n 0)) (entupd blkEnt))
  n)

;; Chama ATTSYNC no bloco para sincronizar instancias com a definicao
(defun STW:sync-bloco (blkname)
  (command "_.ATTSYNC" "_N" blkname)
  (princ (strcat "\n  ATTSYNC: " blkname)))

;; ================================================================
(princ "\nstw-automacao.lsp carregado. Comandos disponiveis:")
(princ "\n  STWAtualizarTags -> atualiza 0E_TAG pai/filho + ID_VISIVEL + NOME_DO_BLOCO")
(princ "\n  STWExportar      -> atualiza tudo e exporta Desktop\\todos_atributos.csv")
(princ "\n  STWImportar      <- importa Desktop\\todos_atributos.csv (Excel -> CAD)")
(princ)
