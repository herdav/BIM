(defun c:setTwoCirclesAndLine (/ p1 p2 ent1 ent2 ent3)
  (setq p1 (getpoint "\nKlicken Sie auf den ersten Punkt: "))
  
  ;; Zeichne den ersten Kreis
  (command "_circle" p1 "d" 100)
  (setq ent1 (entlast)) ; Speichere den Entity-Namen des ersten Kreises
  
  (initget 1)
  
  (while
    (not p2)
    (setq p2 (getpoint "\nKlicken Sie auf den zweiten Punkt: " p1))
    (if p2
      (progn
        (grdraw p1 p2 1 -1) ;; Zeichne Linie für visuelle Darstellung
      )
    )
  )
  
  ;; Zeichne den zweiten Kreis
  (command "_circle" p2 "d" 100)
  (setq ent2 (entlast)) ; Speichere den Entity-Namen des zweiten Kreises
  
  ;; Verbinde die Mittelpunkte mit einer Linie
  (command "_line" p1 p2 "")
  (setq ent3 (entlast)) ; Speichere den Entity-Namen der Linie
  
  ;; Ausgabe der Entity-Namen
  (print (strcat "Entity1: " (vl-princ-to-string ent1)))
  (print (strcat "Entity2: " (vl-princ-to-string ent2)))
  (print (strcat "Entity3: " (vl-princ-to-string ent3)))
)
