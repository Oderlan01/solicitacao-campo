(vl-load-com)

;;; ================================================================
;;; automacao.lsp
;;; Ferramentas AutoCAD para sincronizacao de atributos com Excel
;;;
;;; FLUXO UNIFICADO (arquivo unico):
;;;
;;;   CAD -> Excel:
;;;     1. AutoCAD: ExportarTodos  -> Desktop\todos_atributos.csv
;;;     2. Excel:   ImportarDoCAD  <- Desktop\todos_atributos.csv
;;;
;;;   Excel -> CAD:
;;;     1. Excel:   ExportarParaCAD -> Desktop\todos_atributos.csv
;;;     2. AutoCAD: ImportarTodos   <- Desktop\todos_atributos.csv
;;;
;;; Comandos disponiveis:
;;;   ExportarTodos      -> Desktop\todos_atributos.csv
;;;   ImportarTodos      <- Desktop\todos_atributos.csv
;;;   VerificarAtributos -> relatorio de atributos nos blocos
;;; ================================================================

;; ----------------------------------------------------------------
;; HELPER: remove \r final (CRLF do Windows/Excel)
;; ----------------------------------------------------------------
(defun SC:STRIP-CR (s)
  (if (and (> (strlen s) 0)
           (= (substr s (strlen s) 1) "\r"))
    (substr s 1 (1- (strlen s)))
    s))

;; ----------------------------------------------------------------
;; HELPER: separa string por delimitador (com tratamento CRLF)
;; ----------------------------------------------------------------
(defun quebrar-texto (str delim / pos lst)
  (setq str (SC:STRIP-CR str))
  (while (setq pos (vl-string-search delim str))
    (setq lst (cons (substr str 1 pos) lst))
    (setq str (substr str (+ pos 2))))
  (reverse (cons str lst)))

;; ----------------------------------------------------------------
;; FERRAMENTA 1: EXPORTAR ATRIBUTOS PARA CSV
;; Arquivo: Desktop\todos_atributos.csv
;; Formato: HANDLE_CAD;Nome_Bloco;TAG1;TAG2;...
;; Exporta apenas blocos cujo nome comeca com "ATL_STW".
;; Tambem sincroniza ID_VISIVEL e NOME_DO_BLOCO em cada bloco.
;; ----------------------------------------------------------------
(defun c:ExportarTodos (/ desktop caminho ss i ent obj nomeReal hnd
                           listaAtribs nomeAtrib valorAtrib arq blocosDados
                           listaTags cabecalho linhaTexto busca parAtribs)
  (setq desktop
        (vl-registry-read
          "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders"
          "Desktop"))
  (setq caminho (strcat desktop "\\todos_atributos.csv"))
  (setq ss (ssget "X" '((0 . "INSERT"))))
  (if (not ss)
    (princ "\nNenhum bloco encontrado.")
    (progn
      (setq blocosDados nil listaTags nil i 0)
      (while (< i (sslength ss))
        (setq ent     (ssname ss i)
              obj     (vlax-ename->vla-object ent)
              hnd     (vla-get-Handle obj)
              nomeReal nil)
        (if (vlax-property-available-p obj 'EffectiveName)
          (setq nomeReal (vlax-get-property obj 'EffectiveName)))
        (if (or (null nomeReal) (= nomeReal ""))
          (setq nomeReal (vlax-get-property obj 'Name)))
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
                  (if (= (strcase nomeAtrib) "ID_VISIVEL")
                    (progn (vla-put-TextString atrib hnd) (setq valorAtrib hnd)))
                  (if (= (strcase nomeAtrib) "NOME_DO_BLOCO")
                    (progn (vla-put-TextString atrib nomeReal) (setq valorAtrib nomeReal)))
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
      (command "_regen")
      (princ (strcat "\nExportado: " caminho))))
  (princ))

;; ----------------------------------------------------------------
;; FERRAMENTA 2: IMPORTAR DO CSV (todos_atributos.csv)
;; Lê o arquivo gerado pelo Excel (ExportarParaCAD) ou pelo
;; proprio AutoCAD e atualiza todos os atributos dos blocos.
;; Formato esperado: HANDLE_CAD;Nome_Bloco;TAG1;TAG2;...
;; ----------------------------------------------------------------
(defun c:ImportarTodos (/ desktop caminho arq linha cabecalho listaTags dados
                           hndVal ent obj listaAtribs i tagAtual valorNovo atrib
                           atualizados naoEncontrados)
  (setq desktop
        (vl-registry-read
          "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders"
          "Desktop"))
  (setq caminho (strcat desktop "\\todos_atributos.csv"))
  (setq arq (open caminho "r"))
  (if (not arq)
    (princ (strcat "\nArquivo nao encontrado: " caminho
                   "\nExecute primeiro ExportarParaCAD no Excel."))
    (progn
      (setq linha      (read-line arq)
            cabecalho  (quebrar-texto linha ";")
            listaTags  (cddr cabecalho) ; pula HANDLE_CAD e Nome_Bloco
            atualizados 0
            naoEncontrados 0)
      (while (setq linha (read-line arq))
        (setq dados  (quebrar-texto linha ";")
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
      (command "_regen")
      (princ (strcat "\nImportacao concluida: "
                     (itoa atualizados) " atributo(s) atualizado(s)"
                     (if (> naoEncontrados 0)
                       (strcat ", " (itoa naoEncontrados) " handle(s) nao encontrado(s).")
                       ".")))))
  (princ))

;; ----------------------------------------------------------------
;; FERRAMENTA 3: VERIFICAR ATRIBUTOS NOS BLOCOS
;; Relata quais definicoes de bloco nao possuem os atributos
;; listados em SC:ATRIBS-MANUAIS. Rode antes do BATTMAN+ATTSYNC.
;; ----------------------------------------------------------------
(setq SC:ATRIBS-MANUAIS '("INTERFACE" "CONEXAO" "ORIGEM"))

(defun c:VerificarAtributos (/ ss i ent obj nomeBloco listaAtribs tags
                                vistos faltam faltandoNeste total)
  (setq ss (ssget "X" '((0 . "INSERT"))))
  (if (not ss)
    (princ "\nNenhum bloco encontrado no desenho.")
    (progn
      (setq vistos '() faltam '() i 0)
      (while (< i (sslength ss))
        (setq ent      (ssname ss i)
              obj      (vlax-ename->vla-object ent)
              nomeBloco nil)
        (if (vlax-property-available-p obj 'EffectiveName)
          (setq nomeBloco (vlax-get-property obj 'EffectiveName)))
        (if (or (null nomeBloco) (= nomeBloco ""))
          (setq nomeBloco (vlax-get-property obj 'Name)))
        (if (and nomeBloco (not (member nomeBloco vistos)))
          (progn
            (setq vistos (cons nomeBloco vistos))
            (if (= (vla-get-HasAttributes obj) :vlax-true)
              (progn
                (setq listaAtribs
                      (vlax-safearray->list (vlax-variant-value (vla-GetAttributes obj)))
                      tags '())
                (foreach atrib listaAtribs
                  (setq tags (cons (strcase (vla-get-TagString atrib)) tags)))
                (setq faltandoNeste
                      (vl-remove-if
                        (function (lambda (a) (member (strcase a) tags)))
                        SC:ATRIBS-MANUAIS))
                (if faltandoNeste
                  (setq faltam (cons (list nomeBloco faltandoNeste) faltam)))))))
        (setq i (1+ i)))
      (setq total (length vistos))
      (if (null faltam)
        (princ (strcat "\nOK! Todos os " (itoa total) " tipo(s) de bloco estao corretos."))
        (progn
          (princ "\n=== BLOCOS COM ATRIBUTOS FALTANDO ===")
          (foreach r faltam
            (princ (strcat "\n  Bloco: " (car r)))
            (princ (strcat "\n    Faltam: " (vl-princ-to-string (cadr r)))))
          (princ (strcat "\n\nTotal: " (itoa (length faltam))
                         " definicao(oes) precisam de BATTMAN + ATTSYNC."))))))
  (princ))

;; ----------------------------------------------------------------
(princ "\nautomacao.lsp carregado.")
(princ "\n  ExportarTodos      -> Desktop\\todos_atributos.csv  (CAD -> Excel)")
(princ "\n  ImportarTodos      <- Desktop\\todos_atributos.csv  (Excel -> CAD)")
(princ "\n  VerificarAtributos -> relatorio de atributos faltando")
(princ)
