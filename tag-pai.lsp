(vl-load-com)

;;; ================================================================
;;; tag-pai.lsp
;;; Propaga o VALOR do atributo 0E_TAG do bloco pai para o atributo
;;; 0E_TAG de todos os blocos filhos aninhados dentro dele.
;;;
;;; Resultado: 0E_TAG filho = valor_atual_filho + 0E_TAG_pai
;;; Exemplo:   filho "VP-" + pai "EV1-001"  =>  "VP-EV1-001"
;;;
;;; Fluxo:
;;;   1. Percorre INSERTs no espaco modelo
;;;   2. Para cada pai com 0E_TAG preenchido, modifica filhos na definicao
;;;   3. Usa BEDIT/BCLOSE para confirmar as mudancas e atualizar o desenho
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
;; HELPER: modifica ATTRIBs dos filhos na definicao via entget/entmod.
;;         Retorna lista de nomes de blocos filhos que foram alterados.
;; ----------------------------------------------------------------
(defun TAG-PAI:modificar-def (blkname valorPai / blkRec ent dados tipo
                               subEnt subDados tag valorAtual novoValor
                               modificou n filhosAlterados)
  (setq blkRec        (tblsearch "BLOCK" blkname)
        n             0
        filhosAlterados '())
  (if (not blkRec) (return filhosAlterados))

  (setq ent (cdr (assoc -2 blkRec)))
  (while ent
    (setq dados (entget ent)
          tipo  (cdr (assoc 0 dados)))

    (if (= tipo "INSERT")
      (progn
        (setq nomeFilho (cdr (assoc 2 dados)) ; nome do bloco filho
              subEnt    (entnext ent)
              modificou nil)
        (while (and subEnt
                    (setq subDados (entget subEnt))
                    (= (cdr (assoc 0 subDados)) "ATTRIB"))
          (setq tag (strcase (cdr (assoc 2 subDados))))
          (if (= tag "0E_TAG")
            (progn
              (setq valorAtual (cdr (assoc 1 subDados)))
              (if (= (vl-string-search valorPai valorAtual) nil)
                (progn
                  (setq novoValor (strcat valorAtual valorPai))
                  (entmod (subst (cons 1 novoValor)
                                 (assoc 1 subDados)
                                 subDados))
                  (setq modificou T n (1+ n))))))
          (setq subEnt (entnext subEnt)))
        (if modificou
          (progn
            (entupd ent)
            (if (not (member nomeFilho filhosAlterados))
              (setq filhosAlterados (cons nomeFilho filhosAlterados)))))))

    (setq ent (entnext ent)))

  (list n filhosAlterados))

;; ----------------------------------------------------------------
;; HELPER: usa BEDIT/BCLOSE para confirmar as mudancas da definicao
;;         e propagar para todas as instancias no desenho.
;; ----------------------------------------------------------------
(defun TAG-PAI:commit-bloco (blkname / bloqEdit)
  (setvar "BLOCKEDITLOCK" 0)
  ;; Abre o editor do bloco silenciosamente e fecha salvando
  (command "_.BEDIT" blkname)
  (command "_.BCLOSE" "_S") ; _S = Save
  (princ (strcat "\n  Bloco confirmado: " blkname)))

;; ----------------------------------------------------------------
;; NUCLEO: percorre INSERTs do espaco modelo, le 0E_TAG do pai,
;;         modifica os filhos e confirma via BEDIT/BCLOSE
;; ----------------------------------------------------------------
(defun TAG-PAI:propagar (/ ss i objPai nomePai tagValor
                            resultado n filhos
                            blocosCommit total)
  (setq ss          (ssget "X" '((0 . "INSERT")))
        total       0
        blocosCommit '())
  (if (not ss)
    (progn (princ "\nNenhum bloco encontrado.") (exit)))

  ;; --- PASSO 1: modifica todas as definicoes necessarias ---
  (setq i 0)
  (while (< i (sslength ss))
    (setq objPai  (vlax-ename->vla-object (ssname ss i))
          nomePai (TAG-PAI:nome-efetivo objPai))

    (if (and nomePai (/= (substr nomePai 1 1) "*"))
      (progn
        (setq tagValor (TAG-PAI:get-0e-tag objPai))
        (if (and tagValor (/= tagValor "") (/= tagValor "-"))
          (progn
            (setq resultado (TAG-PAI:modificar-def nomePai tagValor)
                  n         (car resultado)
                  filhos    (cadr resultado))
            (setq total (+ total n))
            ;; Registra bloco pai para commit se houve alteracao
            (if (> n 0)
              (if (not (member nomePai blocosCommit))
                (setq blocosCommit (cons nomePai blocosCommit))))))))
    (setq i (1+ i)))

  ;; --- PASSO 2: confirma cada bloco modificado via BEDIT/BCLOSE ---
  (if blocosCommit
    (progn
      (princ "\nConfirmando alteracoes nos blocos...")
      (foreach blk blocosCommit
        (TAG-PAI:commit-bloco blk))))

  total)

;; ----------------------------------------------------------------
;; COMANDO PRINCIPAL
;; ----------------------------------------------------------------
(defun c:AtualizarTagPai (/ total bloqLock)
  (vl-load-com)
  ;; Garante que o editor de bloco pode ser aberto via comando
  (setq bloqLock (getvar "BLOCKEDITLOCK"))
  (setvar "BLOCKEDITLOCK" 0)

  (setq total (TAG-PAI:propagar))

  (setvar "BLOCKEDITLOCK" bloqLock)
  (command "_.REGENALL")
  (princ (strcat "\nAtualizarTagPai concluido: "
                 (itoa total)
                 " atributo(s) 0E_TAG atualizado(s) nos filhos."))
  (princ))

;; ----------------------------------------------------------------
(princ "\ntag-pai.lsp carregado.")
(princ "\n  AtualizarTagPai  -> propaga 0E_TAG do pai para os filhos aninhados")
(princ)
