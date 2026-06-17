(vl-load-com)

;;; ================================================================
;;; tag-pai.lsp
;;; Propaga o VALOR do atributo 0E_TAG do bloco pai para o atributo
;;; 0E_TAG de todos os blocos filhos aninhados dentro dele.
;;;
;;; Fluxo:
;;;   1. Percorre todos os INSERTs do espaco modelo
;;;   2. Para cada pai com 0E_TAG preenchido, entra na definicao do bloco
;;;   3. Busca INSERTs filhos com 0E_TAG dentro dessa definicao
;;;   4. Escreve o VALOR do 0E_TAG do pai nos filhos
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
;; HELPER: retorna o valor do atributo com a tag informada (strcase)
;;         de um VLA INSERT. Retorna nil se nao encontrar.
;; ----------------------------------------------------------------
(defun TAG-PAI:get-attr (obj tagBusca / atribs val)
  (setq val nil)
  (if (= (vla-get-HasAttributes obj) :vlax-true)
    (progn
      (setq atribs (vlax-safearray->list
                     (vlax-variant-value (vla-GetAttributes obj))))
      (foreach att atribs
        (if (= (strcase (vla-get-TagString att)) (strcase tagBusca))
          (setq val (vla-get-TextString att))))))
  val)

;; ----------------------------------------------------------------
;; HELPER: escreve valor em todos os atributos 0E_TAG de um VLA INSERT
;;         dentro de uma definicao de bloco (entidade aninhada).
;;         Retorna quantidade alterada.
;; ----------------------------------------------------------------
(defun TAG-PAI:set-0e-tag-def (entDef valor / atribs n)
  (setq atribs (vlax-invoke entDef 'GetAttributes)
        n 0)
  (foreach att atribs
    (if (= (strcase (vla-get-TagString att)) "0E_TAG")
      (progn
        (vla-put-TextString att valor)
        (setq n (1+ n)))))
  n)

;; ----------------------------------------------------------------
;; NUCLEO: para cada INSERT pai no espaco modelo,
;;         pega o valor de 0E_TAG e propaga para filhos na definicao
;; ----------------------------------------------------------------
(defun TAG-PAI:propagar (/ doc blocos ss i entPai objPai nomePai
                            tagValor defBloco entFilho total)
  (setq doc    (vla-get-ActiveDocument (vlax-get-acad-object))
        blocos (vla-get-Blocks doc)
        ss     (ssget "X" '((0 . "INSERT")))
        total  0)
  (if (not ss)
    (progn (princ "\nNenhum bloco encontrado.") (exit)))

  (setq i 0)
  (while (< i (sslength ss))
    (setq entPai  (ssname ss i)
          objPai  (vlax-ename->vla-object entPai)
          nomePai (TAG-PAI:nome-efetivo objPai))

    ;; Ignora espacos especiais e blocos sem nome util
    (if (and nomePai (/= (substr nomePai 1 1) "*"))
      (progn
        ;; Pega o valor atual de 0E_TAG deste bloco pai
        (setq tagValor (TAG-PAI:get-attr objPai "0E_TAG"))

        ;; So propaga se o pai tiver 0E_TAG preenchido
        (if (and tagValor (/= tagValor "") (/= tagValor "-"))
          (progn
            ;; Entra na definicao do bloco pai e atualiza filhos
            (setq defBloco (vla-item blocos nomePai))
            (if defBloco
              (vlax-for entFilho defBloco
                (if (and (= (vla-get-ObjectName entFilho) "AcDbBlockReference")
                         (= (vla-get-HasAttributes entFilho) :vlax-true))
                  (setq total
                        (+ total
                           (TAG-PAI:set-0e-tag-def entFilho tagValor))))))))))
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
                 " atributo(s) 0E_TAG propagado(s) dos pais para os filhos."))
  (princ))

;; ----------------------------------------------------------------
(princ "\ntag-pai.lsp carregado.")
(princ "\n  AtualizarTagPai  -> propaga valor de 0E_TAG do pai para os filhos aninhados")
(princ)
