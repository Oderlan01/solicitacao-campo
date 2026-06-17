(vl-load-com)

;;; ================================================================
;;; tag-pai.lsp
;;; Propaga o VALOR do atributo 0E_TAG do bloco pai para o atributo
;;; 0E_TAG de todos os blocos filhos aninhados dentro dele.
;;;
;;; Resultado: 0E_TAG filho = valor_atual_filho + 0E_TAG_pai
;;; Exemplo:   filho "VP-" + pai "EV1-001"  =>  "VP-EV1-001"
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
;;         no espaco modelo. Retorna nil se nao encontrar.
;; ----------------------------------------------------------------
(defun TAG-PAI:get-0e-tag (obj / atribs val)
  (setq val nil)
  (if (= (vla-get-HasAttributes obj) :vlax-true)
    (foreach att (vlax-safearray->list
                   (vlax-variant-value (vla-GetAttributes obj)))
      (if (= (strcase (vla-get-TagString att)) "0E_TAG")
        (setq val (vla-get-TextString att)))))
  val)

;; ----------------------------------------------------------------
;; HELPER: percorre entidades de uma definicao de bloco (via tblsearch
;;         + entnext) e atualiza o 0E_TAG dos INSERTs filhos usando
;;         entget/entmod para garantir persistencia no desenho.
;;         Concatena: valor_atual_filho + valorPai
;;         Retorna quantidade de atributos alterados.
;; ----------------------------------------------------------------
(defun TAG-PAI:atualizar-def (blkname valorPai / blkRec primeiraEnt ent
                                dados tipo subEnt subDados tag
                                valorAtual novoValor n)
  (setq blkRec (tblsearch "BLOCK" blkname)
        n 0)
  (if (not blkRec)
    (progn (princ (strcat "\nDefinicao nao encontrada: " blkname)) (exit)))

  ;; -2 aponta para a primeira entidade da definicao do bloco
  (setq primeiraEnt (cdr (assoc -2 blkRec))
        ent          primeiraEnt)

  (while ent
    (setq dados (entget ent)
          tipo  (cdr (assoc 0 dados)))

    ;; Procura INSERTs aninhados dentro da definicao
    (if (= tipo "INSERT")
      (progn
        ;; Percorre sub-entidades (ATTRIBs) do INSERT filho
        (setq subEnt (entnext ent))
        (while (and subEnt
                    (setq subDados (entget subEnt))
                    (= (cdr (assoc 0 subDados)) "ATTRIB"))
          (setq tag (strcase (cdr (assoc 2 subDados)))) ; grupo 2 = tag
          (if (= tag "0E_TAG")
            (progn
              (setq valorAtual (cdr (assoc 1 subDados))) ; grupo 1 = valor
              ;; Concatena somente se o valorPai ainda nao estiver no resultado
              (if (= (vl-string-search valorPai valorAtual) nil)
                (progn
                  (setq novoValor (strcat valorAtual valorPai))
                  (entmod (subst (cons 1 novoValor)
                                 (assoc 1 subDados)
                                 subDados))
                  (entupd subEnt)
                  (setq n (1+ n))))))
          (setq subEnt (entnext subEnt)))))

    (setq ent (entnext ent)))
  n)

;; ----------------------------------------------------------------
;; NUCLEO: percorre INSERTs do espaco modelo, le 0E_TAG do pai e
;;         propaga para os filhos aninhados na definicao do bloco
;; ----------------------------------------------------------------
(defun TAG-PAI:propagar (/ ss i entPai objPai nomePai tagValor total)
  (setq ss    (ssget "X" '((0 . "INSERT")))
        total 0)
  (if (not ss)
    (progn (princ "\nNenhum bloco encontrado.") (exit)))

  (setq i 0)
  (while (< i (sslength ss))
    (setq entPai  (ssname ss i)
          objPai  (vlax-ename->vla-object entPai)
          nomePai (TAG-PAI:nome-efetivo objPai))

    (if (and nomePai (/= (substr nomePai 1 1) "*"))
      (progn
        (setq tagValor (TAG-PAI:get-0e-tag objPai))
        (if (and tagValor (/= tagValor "") (/= tagValor "-"))
          (setq total (+ total (TAG-PAI:atualizar-def nomePai tagValor))))))

    (setq i (1+ i)))
  total)

;; ----------------------------------------------------------------
;; COMANDO PRINCIPAL
;; ----------------------------------------------------------------
(defun c:AtualizarTagPai (/ total)
  (vl-load-com)
  (setq total (TAG-PAI:propagar))
  (command "_.REGEN")
  (princ (strcat "\nAtualizarTagPai concluido: "
                 (itoa total)
                 " atributo(s) 0E_TAG atualizado(s) nos filhos."))
  (princ))

;; ----------------------------------------------------------------
(princ "\ntag-pai.lsp carregado.")
(princ "\n  AtualizarTagPai  -> propaga 0E_TAG do pai para os filhos aninhados")
(princ)
