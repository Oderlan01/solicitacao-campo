(vl-load-com)

;;; ================================================================
;;; tag-pai.lsp
;;; Propaga o VALOR do atributo 0E_TAG do bloco pai para o atributo
;;; 0E_TAG de todos os blocos filhos aninhados dentro dele.
;;;
;;; Resultado: 0E_TAG filho = ATTDEF_default_filho + 0E_TAG_pai
;;;
;;; Exemplo:
;;;   ATTDEF padrao do filho : "VP-"        (fixo, nunca muda)
;;;   0E_TAG do pai          : "MO-Dois"
;;;   Resultado              : "VP-MO-Dois" (substitui, nao acumula)
;;;
;;; Problema resolvido:
;;;   - Usa o valor PADRAO da ATTDEF do bloco filho como base fixa
;;;   - Chama ATTSYNC no bloco pai para propagar para o modelo
;;;
;;; Comando: AtualizarTagPai
;;; ================================================================

;; ----------------------------------------------------------------
;; HELPER: retorna o nome efetivo de um VLA INSERT (suporta dinamicos)
;; ----------------------------------------------------------------
(defun TAG-PAI:nome-efetivo (obj / nome)
  (setq nome nil)
  (if (vlax-property-available-p obj 'EffectiveName)
    (setq nome (vlax-get-property obj 'EffectiveName)))
  (if (or (null nome) (= nome ""))
    (setq nome (vlax-get-property obj 'Name)))
  nome)

;; ----------------------------------------------------------------
;; HELPER: retorna o valor PADRAO (ATTDEF) do atributo 0E_TAG
;;         dentro da definicao do bloco filho.
;;         Este valor e a "base fixa" que nunca muda.
;; ----------------------------------------------------------------
(defun TAG-PAI:get-attdef-base (childBlkName / blkRec ent dados tipo base)
  (setq blkRec (tblsearch "BLOCK" childBlkName)
        base   "")
  (if blkRec
    (progn
      (setq ent (cdr (assoc -2 blkRec)))
      (while ent
        (setq dados (entget ent)
              tipo  (cdr (assoc 0 dados)))
        ;; ATTDEF = definicao do atributo (grupo 1 = valor padrao, grupo 2 = tag)
        (if (and (= tipo "ATTDEF")
                 (= (strcase (cdr (assoc 2 dados))) "0E_TAG"))
          (setq base (cdr (assoc 1 dados))))
        (setq ent (entnext ent)))))
  base)

;; ----------------------------------------------------------------
;; HELPER: retorna o valor do atributo 0E_TAG de um VLA INSERT
;; ----------------------------------------------------------------
(defun TAG-PAI:get-0e-tag (obj / val)
  (setq val nil)
  (if (= (vla-get-HasAttributes obj) :vlax-true)
    (foreach att (vlax-safearray->list
                   (vlax-variant-value (vla-GetAttributes obj)))
      (if (= (strcase (vla-get-TagString att)) "0E_TAG")
        (setq val (vla-get-TextString att)))))
  val)

;; ----------------------------------------------------------------
;; HELPER: modifica o 0E_TAG dos filhos dentro da definicao do bloco pai
;;         usando base fixa da ATTDEF + valor do pai (sem acumular).
;;         Retorna quantidade de ATTRIBs alterados.
;; ----------------------------------------------------------------
(defun TAG-PAI:modificar-def (blkname valorPai / blkRec ent dados tipo
                               nomeFilho baseFilho novoValor
                               subEnt subDados tag modificou n)
  (setq blkRec (tblsearch "BLOCK" blkname)
        n 0)
  (if (not blkRec) (return n))

  (setq ent (cdr (assoc -2 blkRec)))
  (while ent
    (setq dados (entget ent)
          tipo  (cdr (assoc 0 dados)))

    (if (= tipo "INSERT")
      (progn
        ;; Nome do bloco filho para buscar sua ATTDEF base
        (setq nomeFilho (cdr (assoc 2 dados))
              baseFilho  (TAG-PAI:get-attdef-base nomeFilho)
              novoValor  (strcat baseFilho valorPai)
              subEnt     (entnext ent)
              modificou  nil)

        (while (and subEnt
                    (setq subDados (entget subEnt))
                    (= (cdr (assoc 0 subDados)) "ATTRIB"))
          (setq tag (strcase (cdr (assoc 2 subDados))))
          (if (= tag "0E_TAG")
            (progn
              ;; Substitui sempre com base+pai (nunca acumula)
              (entmod (subst (cons 1 novoValor)
                             (assoc 1 subDados)
                             subDados))
              (setq modificou T n (1+ n))))
          (setq subEnt (entnext subEnt)))

        (if modificou (entupd ent))))

    (setq ent (entnext ent)))
  n)

;; ----------------------------------------------------------------
;; HELPER: chama ATTSYNC no bloco pai para propagar as mudancas
;;         da definicao para todas as instancias no espaco modelo
;; ----------------------------------------------------------------
(defun TAG-PAI:sync-bloco (blkname)
  ;; ATTSYNC nao reseta valores de atributos existentes,
  ;; apenas forca a sincronizacao e refresh das instancias
  (command "_.ATTSYNC" "_N" blkname)
  (princ (strcat "\n  ATTSYNC: " blkname)))

;; ----------------------------------------------------------------
;; NUCLEO: percorre INSERTs do espaco modelo, le 0E_TAG do pai,
;;         modifica a definicao e chama ATTSYNC para atualizar o desenho
;; ----------------------------------------------------------------
(defun TAG-PAI:propagar (/ ss i objPai nomePai tagValor
                            n blocosSync total)
  (setq ss        (ssget "X" '((0 . "INSERT")))
        total     0
        blocosSync '())
  (if (not ss)
    (progn (princ "\nNenhum bloco encontrado.") (exit)))

  ;; --- PASSO 1: modifica as definicoes ---
  (setq i 0)
  (while (< i (sslength ss))
    (setq objPai  (vlax-ename->vla-object (ssname ss i))
          nomePai (TAG-PAI:nome-efetivo objPai))

    (if (and nomePai (/= (substr nomePai 1 1) "*"))
      (progn
        (setq tagValor (TAG-PAI:get-0e-tag objPai))
        (if (and tagValor (/= tagValor "") (/= tagValor "-"))
          (progn
            (setq n (TAG-PAI:modificar-def nomePai tagValor))
            (setq total (+ total n))
            (if (and (> n 0) (not (member nomePai blocosSync)))
              (setq blocosSync (cons nomePai blocosSync)))))))
    (setq i (1+ i)))

  ;; --- PASSO 2: ATTSYNC em cada bloco pai modificado ---
  (if blocosSync
    (progn
      (princ "\nSincronizando instancias no modelo...")
      (foreach blk blocosSync
        (TAG-PAI:sync-bloco blk))))

  total)

;; ----------------------------------------------------------------
;; COMANDO PRINCIPAL
;; ----------------------------------------------------------------
(defun c:AtualizarTagPai (/ total)
  (vl-load-com)
  (setq total (TAG-PAI:propagar))
  (command "_.REGENALL")
  (princ (strcat "\nAtualizarTagPai concluido: "
                 (itoa total)
                 " atributo(s) 0E_TAG atualizado(s) nos filhos."))
  (princ))

;; ----------------------------------------------------------------
(princ "\ntag-pai.lsp carregado.")
(princ "\n  AtualizarTagPai  -> propaga 0E_TAG do pai para os filhos (base ATTDEF + pai)")
(princ)
