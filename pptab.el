;;;
;;; given a region of text, run pptab on it
;;;
(require 'json)
(require 'bbdb-vcard)

(defvar pptab-exe "pptab"
  "A string. Can be used to specify the location of executable -- e.g., ~/bin/pptab.")

(defvar pptab-json-buffer-name "*pptab-json*") ; hold pptab output (JSON)

(defvar pptab-vcard-buffer-name "*pptab-vcard*") ; hold vcard

(defun pptab-json ()
  "Parse the JSON content of the pptab buffer."
  (with-current-buffer pptab-json-buffer-name
    (goto-char 0)
    (json-read)))

(defun pptab-json-to-vcard ()
  "Take the JSON content of the pptab buffer and generate a vcard.
Return as a string."
  (let ((pptab-alist (pptab-json)))
    ;; adapted from bbdb-vcard-from
    (with-temp-buffer
      (let* ((name (alist-get 'name pptab-alist))
	     (structured-name (alist-get 'structuredname pptab-alist))
	     (first-name (alist-get 'first structured-name))
	     (last-name (alist-get 'last structured-name))
					;(aka (bbdb-vcard-bbdb-record-field record 'aka))
					;(organization (bbdb-vcard-bbdb-record-field record 'organization))
             (emails (alist-get 'emails pptab-alist))
             (phones (alist-get 'phones pptab-alist))
             (address (alist-get 'address pptab-alist))
	     (uri (alist-get 'uri pptab-alist))
             (notes (alist-get 'notes pptab-alist))
             ;; (raw-anniversaries
             ;;  (let ((anniversary (bbdb-vcard-bbdb-record-field record 'anniversary)))
	     ;; 	(when anniversary
	     ;; 	  (bbdb-vcard-split-structured-text anniversary "\n" t))))
             ;; (birthday-regexp
             ;;  "\\([0-9]\\{4\\}-[01][0-9]-[0-3][0-9][t:0-9]*[-+z:0-9]*\\)\\([[:blank:]]+birthday\\)?\\'")
             ;; (birthday
             ;;  (when raw-anniversaries
	     ;; 	(car (bbdb-vcard-split-structured-text
             ;;          (cl-find-if (lambda (x) (string-match birthday-regexp x))
	     ;; 			  raw-anniversaries)
             ;;          " " t))))
             ;; (other-anniversaries
             ;;  (when raw-anniversaries
	     ;; 	(cl-remove-if (lambda (x) (string-match birthday-regexp x))
             ;;                  raw-anniversaries :count 1)))
             ;; (timestamp (bbdb-vcard-bbdb-record-field record 'timestamp))
             ;; (mail-aliases (bbdb-vcard-bbdb-record-field record 'mail-alias))
             ;; (raw-notes (copy-alist (bbdb-record-xfields record)))
	     )
	(bbdb-vcard-insert-vcard-element "BEGIN" "VCARD")
	(bbdb-vcard-insert-vcard-element "VERSION" "3.0")
	(when name
	  (bbdb-vcard-insert-vcard-element "FN" (bbdb-vcard-escape-strings name)))
	;; Per vcard specification(s), the N element is not optional in a vcard.
	(cond ((or last-name first-name)
	       (bbdb-vcard-insert-vcard-element
		"N" (bbdb-vcard-escape-strings last-name)
		";" (bbdb-vcard-escape-strings first-name)
		";;;" ; Additional Names, Honorific Prefixes, Honorific Suffixes
		))
	      (t
	       (display-warning 'pptab
				"N field is required" 
				:error))) 
	;; (when aka
	;;   (bbdb-vcard-insert-vcard-element
	;;    "NICKNAME" (bbdb-join (bbdb-vcard-escape-strings aka) ",")))
	;; (dolist (org organization)
	;;   (bbdb-vcard-insert-vcard-element
	;;    "ORG" (bbdb-vcard-escape-strings org)))
	;; (dolist (mail net)
	;;   (bbdb-vcard-insert-vcard-element
	;;    "EMAIL;TYPE=INTERNET" (bbdb-vcard-escape-strings mail)))
	(let ((emails-length (length emails)))
	  (dotimes (i emails-length)
	    (let ((email (elt emails i)))
	      (bbdb-vcard-insert-vcard-element
	       "EMAIL;TYPE=INTERNET"
	       (bbdb-vcard-escape-strings email)))))
	(let ((phones-length (length phones)))
	  (dotimes (i phones-length)
	    (let ((phone (elt phones i)))
	      (bbdb-vcard-insert-vcard-element
	       (concat
		"TEL;TYPE="
		;; no support for label
		;; (bbdb-vcard-escape-strings
		;; 	(bbdb-vcard-translate (alist-get bbdb-phone-label phone) t))
		)
	       (bbdb-vcard-escape-strings phone)))))
	;; address should be a sequence
	(when address
	  (pptab-vcard-insert-address
	   (append (elt (alist-get 'address (pptab-json)) 0) (elt (alist-get 'address (pptab-json)) 1))))
	;; (let ((address-length (length address)))
	;;   (dotimes (i address-length)
	;;     (let ((address-line (elt address i)))
	;;       (when address-line
	;; 	))))
	(when uri
	  ;; uri is a vector with structure ["https", "www.gizdich-ranch.com", "/u-pick", "", "", ""]
	  (bbdb-vcard-insert-vcard-element "URL"
					   ;; assume http and https uris - turn to a library for anything more substantial
					   (s-join nil
						   (vconcat (vector (elt uri 0) "://")
							    (subseq uri 1)))))
	(when notes
	  (let ((notes-length (length notes)))
	    (dotimes (i notes-length)
	      (let ((note (elt notes i)))
		(bbdb-vcard-insert-vcard-element "NOTE"
						 (bbdb-vcard-escape-strings note))))))
	;; (when birthday
	;;   (bbdb-vcard-insert-vcard-element "BDAY" birthday))
	;; (when other-anniversaries
	;;   (bbdb-vcard-insert-vcard-element  ; non-birthday anniversaries
	;;    "X-BBDB-ANNIVERSARY" (bbdb-join other-anniversaries "\\n")))
	;; (when timestamp
	;;   (bbdb-vcard-insert-vcard-element "REV" timestamp))
	;; (when mail-aliases
	;;   (bbdb-vcard-insert-vcard-element
	;;    "CATEGORIES"
	;;    (bbdb-join (bbdb-vcard-escape-strings
	;;                (bbdb-vcard-split-structured-text mail-aliases "," t)) ",")))
	;; ;; If fields have been export, prune from raw-notes ...
	;; (dolist (key `(url notes anniversary mail-alias creation-date timestamp))
	;;   (setq raw-notes (assq-delete-all key raw-notes)))
	;; ;; ... and output what's left
	;; (dolist (raw-note raw-notes)
	;;   (when raw-note
	;;     (bbdb-vcard-insert-vcard-element
	;;      (symbol-name (bbdb-vcard-prepend-x-bbdb-maybe (car raw-note)))
	;;      (bbdb-vcard-escape-strings (cdr raw-note)))))
	(bbdb-vcard-insert-vcard-element "END" "VCARD")
	(bbdb-vcard-insert-vcard-element nil)) ; newline
      (buffer-string))
    ;; END -- adapted from bbdb-vcard-from
    ))

(defun pptab-replace ()
  "Run pptab on a region of text. Replace the region with the
corresponding vcard."
  (interactive)
  (let ((point (point))
	(mark (mark)))
    (pptab-to-vcard)			; run pptab on region
    (delete-region point mark))
  (let ((vcard-string (with-current-buffer pptab-vcard-buffer-name
			(buffer-string))))
    (insert vcard-string)))

(defun pptab-string-to-json (string)
  "Run pptab on STRING. Send the JSON output to
PPTAB-JSON-BUFFER-NAME. Return value undefined."
  ;; prepare pptab-json-buffer
  (get-buffer-create pptab-json-buffer-name)
  (with-current-buffer pptab-json-buffer-name
    (delete-region (point-min) (point-max)))
  ;; prepare input file for pptab
  (let ((file-name (make-temp-file "foo")))
    (with-temp-buffer
      (insert string)
      (write-region (point-min)
                    (point-max)
                    file-name))
    ;; invoke pptab
    (call-process pptab-exe
		  nil	  ; program's input comes from file file-name 
		  pptab-json-buffer-name
		  t
		  file-name)
    (delete-file file-name)))

(defun pptab-to-json ()
  "Run pptab on the selected region of text. Send the JSON output to
PPTAB-JSON-BUFFER-NAME."
  (interactive)
  (get-buffer-create pptab-json-buffer-name)
  (with-current-buffer pptab-json-buffer-name
    (delete-region (point-min) (point-max)))
  (let ((point (point))
	(mark (mark)))
    (let ((file-name (make-temp-file "foo")))
      (write-region point mark file-name)
      (call-process pptab-exe
		    nil	  ; program's input comes from file file-name 
		    pptab-json-buffer-name
		    t
		    file-name)
      (delete-file file-name))))

(defun pptab-to-vcard (&optional print-message)
  "Run pptab on the selected region of text. View the corresponding
vcard in the pptab-vcard (see PPTAB-VCARD-BUFFER-NAME) buffer."
  (interactive "p")
  (pptab-to-json)
  (get-buffer-create pptab-vcard-buffer-name)
  (with-current-buffer pptab-vcard-buffer-name
    (delete-region (point-min) (point-max))
    (insert (pptab-json-to-vcard)))
  ;; Typically, interactive use will include an interest in viewing the vcard
  (when print-message
    (switch-to-buffer pptab-vcard-buffer-name)))

(defun pptab-vcard-insert-address (address)
  "ADDRESS should be an alist."
  (let ((place-name (alist-get 'PlaceName address))
	(state-name (alist-get 'StateName address))
	(zip-code (alist-get 'ZipCode address)))
    (bbdb-vcard-insert-vcard-element
     (concat
      "ADR;TYPE="
      ;; No support for address labels
      ;; (bbdb-vcard-escape-strings
      ;;  (bbdb-vcard-translate (bbdb-address-label address) t))
      )
     ";;"			; no Postbox, no Extended
     (bbdb-join (list
		 (alist-get 'AddressNumber address)
		 (alist-get 'StreetName address)
		 (alist-get 'StreetNamePostType address))
		" ")
     ";" (when place-name (bbdb-vcard-vcardize-address-element
			   (bbdb-vcard-escape-strings (alist-get 'PlaceName address))))
     ";" (when state-name (bbdb-vcard-vcardize-address-element
			   (bbdb-vcard-escape-strings (alist-get 'StateName address))))
     ";" (when zip-code (bbdb-vcard-vcardize-address-element
			 (bbdb-vcard-escape-strings (alist-get 'ZipCode address))))
     ";" (when (alist-get 'Country ; ??
			  address)
	   (bbdb-vcard-vcardize-address-element
	    (bbdb-vcard-escape-strings (bbdb-address-country address)))))))
