;; 块表格文本提取脚本
(defun c:BlockTableToExcel (/ ss blk ent txt data filename)
  ;; 选择块表格
  (setq ss (ssget '((0 . "INSERT"))))
  (if ss
    (progn
      ;; 初始化数据列表
      (setq data '())
      
      ;; 遍历选中的块
      (setq i 0)
      (repeat (sslength ss)
        (setq blk (ssname ss i))
        (setq ent (entnext blk))
        
        ;; 遍历块内实体
        (while ent
          (if (= (cdr (assoc 0 (entget ent))) "TEXT")
            (setq data (append data (list (cdr (assoc 1 (entget ent))))))
          )
          (setq ent (entnext ent))
        )
        (setq i (1+ i))
      )
      
      ;; 导出数据到CSV文件
      (setq filename (getfiled "保存为CSV" "" "csv" 8))
      (if filename
        (progn
          (setq file (open filename "w"))
          (foreach txt data
            (write-line txt file)
          )
          (close file)
          (alert (strcat "导出完成: " filename))
        )
      )
    )
    (alert "未选择任何块!")
  )
  (princ)
)