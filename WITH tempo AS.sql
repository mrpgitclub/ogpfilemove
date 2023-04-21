WITH tempo AS
(
	SELECT  top 100 mqp.moldrite_qa_pull_id
	       ,mqp.part_id
	       ,COUNT(mqpb.object_id) AS Boxes
	FROM moldrite_qa_pull AS mqp
	LEFT OUTER JOIN part
	ON mqp.part_id = part.part_id
	LEFT OUTER JOIN moldrite_qa_pull_boxes AS mqpb
	ON mqp.moldrite_qa_pull_id = mqpb.moldrite_qa_pull_id
	GROUP BY  mqp.moldrite_qa_pull_id
	         ,mqp.part_id
	ORDER BY mqp.moldrite_qa_pull_id DESC
), tempb AS
(
	SELECT  tempo.moldrite_qa_pull_id
	       ,mqpb.object_id
	       ,HIS.quantity * part.current_cost AS CostPerBox
	FROM tempo, moldrite_qa_pull_boxes AS mqpb, object_history AS HIS, part
	WHERE mqpb.moldrite_qa_pull_id = tempo.moldrite_qa_pull_id
	AND HIS.operation_code = 'D'
	AND HIS.object_id = mqpb.object_id
	AND HIS.part_id = part.part_id 
)
SELECT  tempb.moldrite_qa_pull_id
       ,SUM(tempb.CostPerBox) AS Costing
FROM tempb
GROUP BY  tempb.moldrite_qa_pull_id
ORDER BY tempb.moldrite_qa_pull_id DESC