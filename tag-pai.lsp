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
;; NUCLEO: modifica 0E_TAG nos filhos aninhados dentro da definicao
;;         do bloco pai usando entget/entmod/entupd.
;;         Apos modificar a definicao, forca entupd em todas as
;;         instancias do bloco pai no espaco modelo para atualizar
;;         a visualizacao no desenho.
;;         Retorna quantidade de atributos alterados.
;; ----------------------------------------------------------------
(defun TAG-PAI:atualizar-def (blkname valorPai / blkRec ent dados tipo
                                subEnt subDados tag valorAtual novoValor
                                modificou n ssInst j)
  (setq blkRec (tblsearch "BLOCK" blkname)
        n 0)
  (if (not blkRec) (exit))

  ;; --- PASSO A: modifica ATTRIBs dos filhos na definicao do bloco ---
  (setq ent (cdr (assoc -2 blkRec)))
  (while ent
    (setq dados (entget ent)
          tipo  (cdr (assoc 0 dados)))

    (if (= tipo "INSERT")
      (progn
        (setq subEnt   (entnext ent)
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
        ;; entupd no INSERT filho dentro da definicao
        (if modificou (entupd ent))))

    (setq ent (entnext ent)))

  ;; --- PASSO B: forca entupd em todas as instancias do pai no modelo ---
  ;; Isso faz o AutoCAD re-renderizar as instancias com os novos valores
  (if (> n 0)
    (progn
      (setq ssInst (ssget "X" (list (cons 0 "INSERT") (cons 2 blkname))))
      (if ssInst
        (progn
          (setq j 0)
          (while (< j (sslength ssInst))
            (entupd (ssname ssInst j))
            (setq j (1+ j)))))))
  n)

;; ----------------------------------------------------------------
;; NUCLEO: percorre INSERTs do espaco modelo, le 0E_TAG do pai e
;;         delega a atualizacao dos filhos
;; ----------------------------------------------------------------
(defun TAG-PAI:propagar (/ ss i objPai nomePai tagValor total)
  (setq ss    (ssget "X" '((0 . "INSERT")))
        total 0)
  (if (not ss)
    (progn (princ "\nNenhum bloco encontrado.") (exit)))

  (setq i 0)
  (while (< i (sslength ss))
    (setq objPai  (vlax-ename->vla-object (ssname ss i))
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
  (command "_.REGENALL")
  (princ (strcat "\nAtualizarTagPai concluido: "
                 (itoa total)
                 " atributo(s) 0E_TAG atualizado(s) nos filhos."))
  (princ))

;; ----------------------------------------------------------------
(princ "\ntag-pai.lsp carregado.")
(princ "\n  AtualizarTagPai  -> propaga 0E_TAG do pai para os filhos aninhados")
(princ)
