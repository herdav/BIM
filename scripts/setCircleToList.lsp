(defun c:ExportSimilarCircles ()
  ;; Ein Kreis auswählen
  (setq selection (entsel "\nWählen Sie einen Kreis aus: "))
  
  ;; Wenn keine Auswahl getroffen wurde, beenden
  (if (not selection)
    (exit)
  )

  ;; Daten des ausgewählten Kreises abrufen
  (setq selectedCircle (entget (car selection)))
  
  ;; Radius und Layer des ausgewählten Kreises erfassen
  (setq selectedRadius (cdr (assoc 40 selectedCircle)))
  (setq selectedLayer (cdr (assoc 8 selectedCircle)))

  ;; Alle Kreise in der Zeichnung erfassen
  (setq allCircles (ssget "X" '((0 . "CIRCLE"))))
  
  ;; Leere Liste für ähnliche Kreise
  (setq similarCircles nil)
  
  ;; Jeden Kreis überprüfen
  (if allCircles
    (progn
      (setq count (sslength allCircles))
      (repeat count
        (setq ent (ssname allCircles (setq count (- count 1))))
        (setq data (entget ent))
        (if (and (= (cdr (assoc 40 data)) selectedRadius) 
                 (= (cdr (assoc 8 data)) selectedLayer))
          (setq similarCircles (cons data similarCircles))
        )
      )
    )
  )

  ;; Wenn keine ähnlichen Kreise gefunden wurden, beenden
  (if (null similarCircles)
    (progn
      (alert "Keine ähnlichen Kreise gefunden.")
      (exit)
    )
  )
  
  ;; Exportieren
  (setq filename (getfiled "Speichern unter" "" "txt" 1))
  
  (if filename
    (progn
      (setq file (open filename "w"))

      (foreach circle similarCircles
        (setq center (cdr (assoc 10 circle)))
        (write-line (strcat "Kreis auf Layer: " selectedLayer " | Radius: " (rtos selectedRadius) " | Koordinaten: " 
                   "(" (rtos (car center)) ", " (rtos (cadr center)) ", " (rtos (caddr center)) ")") file)
      )

      (close file)
      (alert (strcat "Daten wurden nach " filename " exportiert."))
    )
    (alert "Export abgebrochen.")
  )
)
