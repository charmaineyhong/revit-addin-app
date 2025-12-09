using System;

namespace PPVCRevitAddin
{
    /// <summary>
    /// Represents one row from predictions.csv for a node.
    /// Stores the prediction from the single 4-class model:
    /// - no_annotation (0)
    /// - dimension (1)
    /// - text (2)
    /// - both (3)
    /// </summary>
    public class PredictionRecord
    {
        public long NodeId;
        public int PredictedClass;
        public double Confidence;
        public string AnnotationType;

        public bool NeedDimension => PredictedClass == 1 || PredictedClass == 3;
        public bool NeedText => PredictedClass == 2 || PredictedClass == 3;
    }
}
