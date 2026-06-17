(vl-load-com)

;;; ================================================================
;;; tag-pai.lsp
;;; Preenche o atributo 0E_TAG dos blocos filhos (aninhados) com o
;;; nome do bloco pai que os contem na definicao.
;;;
;;; Comando disponivel:
;;;   AtualizarTagPai  -> atualiza 0E_TAG em todos os filhos aninhados
;;;
;;; Regras:
;;;   - Bloco filho (INSERT dentro de uma definicao de bloco):
;;;       0E_TAG = nome da definicao pai
;;;   - Bloco raiz (INSERT direto no espaco modelo sem pai):
;;;       0E_TAG vazio permanece "-"
;;; ================================================================

;; ----------------------------------------------------------------
;; HELPER: retorna o nome efetivo de um objeto INSERT (VLA)
;; Suporta blocos dinamicos que expõem EffectiveName
;; ----------------------------------------------------------------
(defun TAG-PAI:nome-efetivo (obj / nome)
  (setq nome nil)
  (if (vlax-property-available-p obj 'EffectiveName)
    (setq nome (vlax-get-property obj 'EffectiveName)))
  (if (or (null nome) (= nome ""))
    (setq nome (vlax-get-property obj 'Name)))
  nome)

;; ----------------------------------------------------------------
;; HELPER: percorre atributos de um VLA-object INSERT e atualiza
;; 0E_TAG com o valor fornecido. Retorna quantidade de tags alteradas.
;; ----------------------------------------------------------------
(defun TAG-PAI:set-0e-tag (obj valor / atribs n)
  (setq atribs (vlax-safearray->list
                 (vlax-variant-value (vla-GetAttributes obj)))
        n 0)
  (foreach att atribs
    (if (= (strcase (vla-get-TagString att)) "0E_TAG")
      (progn
        (vla-put-TextString att valor)
        (setq n (1+ n)))))
  n)

;; ----------------------------------------------------------------
;; PASSO 1: Blocos aninhados dentro de definicoes de bloco
;;   Para cada definicao de bloco pai, busca INSERTs filhos com
;;   0E_TAG e preenche com o nome da definicao pai.
;; ----------------------------------------------------------------
(defun TAG-PAI:atualizar-filhos (/ doc blocos nomePai defBloco ent total)
  (setq doc    (vla-get-ActiveDocument (vlax-get-acad-object))
        blocos (vla-get-Blocks doc)
        total  0)
  (vlax-for defBloco blocos
    (if (and (= (vla-get-IsLayout defBloco) :vlax-false)
             (= (vla-get-IsXRef   defBloco) :vlax-false)
             (/= (substr (vla-get-Name defBloco) 1 1) "*"))
      (progn
        (setq nomePai (vla-get-Name defBloco))
        (vlax-for ent defBloco
          (if (and (= (vla-get-ObjectName ent) "AcDbBlockReference")
                   (= (vla-get-HasAttributes ent) :vlax-true))
            (setq total (+ total (TAG-PAI:set-0e-tag ent nomePai))))))))
  total)

;; ----------------------------------------------------------------
;; PASSO 2: Blocos raiz no Espaco Modelo
;;   Blocos inseridos diretamente no espaco modelo (sem pai) que
;;   tenham 0E_TAG vazio recebem "-" para sinalizar nivel raiz.
;; ----------------------------------------------------------------
(defun TAG-PAI:marcar-raiz (/ ss i ent obj nome atribs att)
  (setq ss (ssget "X" '((0 . "INSERT"))))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i)
              obj (vlax-ename->vla-object ent)
              nome (TAG-PAI:nome-efetivo obj))
        (if (and nome
                 (/= (substr nome 1 1) "*")
                 (= (vla-get-HasAttributes obj) :vlax-true))
          (progn
            (setq atribs (vlax-safearray->list
                           (vlax-variant-value (vla-GetAttributes obj))))
            (foreach att atribs
              (if (and (= (strcase (vla-get-TagString att)) "0E_TAG")
                       (= (vla-get-TextString att) ""))
                (vla-put-TextString att "-")))))
        (setq i (1+ i))))))

;; ----------------------------------------------------------------
;; COMANDO PRINCIPAL: AtualizarTagPai
;; ----------------------------------------------------------------
(defun c:AtualizarTagPai (/ total)
  (vl-load-com)
  (setq total (TAG-PAI:atualizar-filhos))
  (TAG-PAI:marcar-raiz)
  (command "_.REGEN")
  (princ (strcat "\nAtualizarTagPai concluido: "
                 (itoa total)
                 " atributo(s) 0E_TAG atualizado(s) em blocos filhos."))
  (princ))

;; ----------------------------------------------------------------
(princ "\ntag-pai.lsp carregado.")
(princ "\n  AtualizarTagPai  -> preenche 0E_TAG dos filhos com nome do bloco pai")
(princ)
