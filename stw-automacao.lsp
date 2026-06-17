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

;; Retorna o nome efetivo de um VLA INSERT (suporta dinamicos)
(defun STW:nome-efetivo (obj / nome)
  (setq nome nil)
  (if (vlax-property-available-p obj 'EffectiveName)
    (setq nome (vlax-get-property obj 'EffectiveName)))
  (if (or (null nome) (= nome ""))
    (setq nome (vlax-get-property obj 'Name)))
  nome)

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
        (if (and nomeReal (/= (substr nomeReal 1 1) "*"))
          (progn
            (setq tagValor (STW:get-0e-tag obj))
            (if (and tagValor (/= tagValor "") (/= tagValor "-"))
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
        (if (and (= (strcase (substr nomeReal 1 7)) "ATL_STW")
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
        (if (and (= (strcase (substr nomeReal 1 7)) "ATL_STW")
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
;; Resultado: 0E_TAG filho = ATTDEF_base_filho + 0E_TAG_pai
;; Exemplo:   base "VP-" + pai "MO-Dois" => "VP-MO-Dois"
;; (nunca acumula — a base vem sempre da ATTDEF do filho)
;; ================================================================

;; Retorna o valor padrao (ATTDEF) do atributo 0E_TAG do bloco filho.
;; Esse valor e a "base fixa" que nunca muda entre atualizacoes.
(defun STW:get-attdef-base (childBlkName / blkRec ent dados tipo base)
  (setq blkRec (tblsearch "BLOCK" childBlkName)
        base   "")
  (if blkRec
    (progn
      (setq ent (cdr (assoc -2 blkRec)))
      (while ent
        (setq dados (entget ent)
              tipo  (cdr (assoc 0 dados)))
        (if (and (= tipo "ATTDEF")
                 (= (strcase (cdr (assoc 2 dados))) "0E_TAG"))
          (setq base (cdr (assoc 1 dados))))
        (setq ent (entnext ent)))))
  base)

;; Retorna o valor do atributo 0E_TAG de um VLA INSERT
(defun STW:get-0e-tag (obj / val)
  (setq val nil)
  (if (= (vla-get-HasAttributes obj) :vlax-true)
    (foreach att (vlax-safearray->list
                   (vlax-variant-value (vla-GetAttributes obj)))
      (if (= (strcase (vla-get-TagString att)) "0E_TAG")
        (setq val (vla-get-TextString att)))))
  val)

;; Modifica 0E_TAG dos filhos na definicao do bloco pai
;; usando entget/entmod/entupd para persistencia no desenho.
;; Usa base fixa da ATTDEF + valor do pai (sem acumular).
(defun STW:modificar-def (blkname valorPai / blkRec ent dados tipo
                           nomeFilho baseFilho novoValor
                           subEnt subDados tag modificou n)
  (setq blkRec (tblsearch "BLOCK" blkname) n 0)
  (if (not blkRec) (return n))
  (setq ent (cdr (assoc -2 blkRec)))
  (while ent
    (setq dados (entget ent)
          tipo  (cdr (assoc 0 dados)))
    (if (= tipo "INSERT")
      (progn
        (setq nomeFilho (cdr (assoc 2 dados))
              baseFilho  (STW:get-attdef-base nomeFilho)
              novoValor  (strcat baseFilho valorPai)
              subEnt     (entnext ent)
              modificou  nil)
        (while (and subEnt
                    (setq subDados (entget subEnt))
                    (= (cdr (assoc 0 subDados)) "ATTRIB"))
          (if (= (strcase (cdr (assoc 2 subDados))) "0E_TAG")
            (progn
              (entmod (subst (cons 1 novoValor)
                             (assoc 1 subDados)
                             subDados))
              (setq modificou T n (1+ n))))
          (setq subEnt (entnext subEnt)))
        (if modificou (entupd ent))))
    (setq ent (entnext ent)))
  n)

;; Chama ATTSYNC no bloco pai para propagar mudancas para o modelo
(defun STW:sync-bloco (blkname)
  (command "_.ATTSYNC" "_N" blkname)
  (princ (strcat "\n  ATTSYNC: " blkname)))

;; ================================================================
(princ "\nstw-automacao.lsp carregado. Comandos disponiveis:")
(princ "\n  STWAtualizarTags -> atualiza 0E_TAG pai/filho + ID_VISIVEL + NOME_DO_BLOCO")
(princ "\n  STWExportar      -> atualiza tudo e exporta Desktop\\todos_atributos.csv")
(princ "\n  STWImportar      <- importa Desktop\\todos_atributos.csv (Excel -> CAD)")
(princ)
