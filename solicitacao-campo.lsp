;;; solicitacao-campo.lsp
;;; Sincronizacao de atributos de blocos AutoCAD com CSV
;;; Uso: SC-EXPORTAR | SC-IMPORTAR
;;;
;;; Formato CSV: HANDLE,TAG,VALOR
;;;   HANDLE  = handle unico do bloco (coluna 5 do DXF)
;;;   TAG     = nome do atributo (group code 2 da entidade ATTRIB)
;;;   VALOR   = conteudo do atributo (group code 1)
;;;
;;; O arquivo CSV pode ser aberto/editado no Excel e reimportado.

;; ----------------------------------------------------------------
;; HELPER: separa linha CSV em lista de campos (delimitador virgula)
;; ----------------------------------------------------------------
(defun SC:LER-LINHA-CSV (linha / resultado pos c acum)
  (setq resultado '()
        pos       0
        acum      "")
  (while (< pos (strlen linha))
    (setq pos (1+ pos)
          c   (substr linha pos 1))
    (if (= c ",")
      (setq resultado (append resultado (list acum))
            acum      "")
      (setq acum (strcat acum c))))
  (append resultado (list acum)))

;; ----------------------------------------------------------------
;; HELPER: remove \r final para compatibilidade com CRLF do Windows
;; ----------------------------------------------------------------
(defun SC:STRIP-CR (s)
  (if (and (> (strlen s) 0)
           (= (substr s (strlen s) 1) "\r"))
    (substr s 1 (1- (strlen s)))
    s))

;; ----------------------------------------------------------------
;; HELPER: retorna lista de (TAG . VALOR) dos atributos de um INSERT
;; ----------------------------------------------------------------
(defun SC:OBTER-ATRIBS (ins-ename / sub-ename sub-edata atribs)
  (setq atribs    '()
        sub-ename (entnext ins-ename))
  (while (and sub-ename
              (= "ATTRIB"
                 (cdr (assoc 0 (setq sub-edata (entget sub-ename))))))
    (setq atribs    (append atribs
                            (list (cons (cdr (assoc 2 sub-edata))
                                        (cdr (assoc 1 sub-edata)))))
          sub-ename (entnext sub-ename)))
  atribs)

;; ----------------------------------------------------------------
;; EXPORTAR: varre todos os blocos com atributos e grava CSV
;; ----------------------------------------------------------------
(defun SC:EXPORTAR-ATRIBUTOS
    (/ arquivo caminho ss idx ename edata handle atribs par total)
  (setq caminho (getfiled "Salvar CSV de Atributos" "" "csv" 1))
  (if (not caminho)
    (princ "\nExportacao cancelada.")
    (progn
      (setq arquivo (open caminho "w"))
      (if (not arquivo)
        (princ "\nErro: nao foi possivel criar o arquivo.")
        (progn
          (write-line "HANDLE,TAG,VALOR" arquivo)
          (setq ss    (ssget "X" '((0 . "INSERT")))
                total 0)
          (if ss
            (progn
              (setq idx 0)
              (while (< idx (sslength ss))
                (setq ename  (ssname ss idx)
                      edata  (entget ename)
                      handle (cdr (assoc 5 edata)))
                (if (= (cdr (assoc 66 edata)) 1)
                  (progn
                    (setq atribs (SC:OBTER-ATRIBS ename))
                    (foreach par atribs
                      (write-line
                        (strcat handle "," (car par) "," (cdr par))
                        arquivo)
                      (setq total (1+ total)))))
                (setq idx (1+ idx)))))
          (close arquivo)
          (princ (strcat "\nExportados " (itoa total)
                         " atributo(s) para:\n" caminho))))))
  (princ))

;; ----------------------------------------------------------------
;; IMPORTAR: le CSV e atualiza atributos dos blocos pelo handle
;; ----------------------------------------------------------------
(defun SC:IMPORTAR-ATRIBUTOS
    (/ arquivo caminho linha campos handle tag valor
       ins-ename sub-ename sub-edata atualizados erros linha-num)
  (setq caminho (getfiled "Abrir CSV de Atributos" "" "csv" 4))
  (if (not caminho)
    (princ "\nImportacao cancelada.")
    (progn
      (setq arquivo (open caminho "r"))
      (if (not arquivo)
        (princ "\nErro: nao foi possivel abrir o arquivo.")
        (progn
          (read-line arquivo) ; pula cabecalho HANDLE,TAG,VALOR
          (setq atualizados 0
                erros       0
                linha-num   1)
          (while (setq linha (read-line arquivo))
            (setq linha-num (1+ linha-num)
                  linha     (SC:STRIP-CR linha)
                  campos    (SC:LER-LINHA-CSV linha))
            (if (< (length campos) 3)
              (progn
                (princ (strcat "\nLinha " (itoa linha-num)
                               " ignorada (formato invalido): " linha))
                (setq erros (1+ erros)))
              (progn
                (setq handle (nth 0 campos)
                      tag    (strcase (nth 1 campos))
                      valor  (nth 2 campos))
                (setq ins-ename (handent handle))
                (if (not ins-ename)
                  (progn
                    (princ (strcat "\nHandle nao encontrado: " handle))
                    (setq erros (1+ erros)))
                  (progn
                    ;; percorre atributos do bloco ate encontrar a TAG
                    (setq sub-ename (entnext ins-ename)
                          sub-edata nil)
                    (while (and sub-ename
                                (= "ATTRIB"
                                   (cdr (assoc 0
                                          (setq sub-edata (entget sub-ename)))))
                                (not (= tag
                                        (strcase (cdr (assoc 2 sub-edata))))))
                      (setq sub-ename (entnext sub-ename)))
                    (if (and sub-ename
                             sub-edata
                             (= "ATTRIB" (cdr (assoc 0 sub-edata)))
                             (= tag (strcase (cdr (assoc 2 sub-edata)))))
                      (progn
                        (entmod (subst (cons 1 valor)
                                       (assoc 1 sub-edata)
                                       sub-edata))
                        (entupd sub-ename)
                        (setq atualizados (1+ atualizados)))
                      (progn
                        (princ (strcat "\nTag '" tag
                                       "' nao encontrada no bloco "
                                       handle))
                        (setq erros (1+ erros)))))))))
          (close arquivo)
          (princ (strcat "\nImportacao concluida: "
                         (itoa atualizados) " atributo(s) atualizado(s)"
                         (if (> erros 0)
                           (strcat ", " (itoa erros)
                                   " aviso(s) — veja o Command Prompt.")
                           ".")))))))
  (princ))

;; ----------------------------------------------------------------
;; Comandos AutoCAD
;; ----------------------------------------------------------------
(defun C:SC-EXPORTAR () (SC:EXPORTAR-ATRIBUTOS))
(defun C:SC-IMPORTAR () (SC:IMPORTAR-ATRIBUTOS))

(princ "\nSolicitacao-Campo carregado.")
(princ "\n  SC-EXPORTAR  ->  exporta atributos dos blocos para CSV")
(princ "\n  SC-IMPORTAR  ->  atualiza atributos a partir de CSV")
(princ)
