public class Plan
{
    public string LessonCode { get; set; }
    public string YearLevel { get; set; }
    public string Duration { get; set; }
    public string KlaTopic { get; set; }
    public string SchoolCode { get; set; }
    public string Curriculum { get; set; }
    public List<Learner> Learners { get; set; }
    public List<Resource> Resources { get; set; }
    public string PedagogicalStrategy { get; set; }
    public string ContentDescription { get; set; }
    public List<Elaboration> Elaborations { get; set; }
    public List<Outcome> LearningOutcomes { get; set; }
    public List<Success> SuccessCriteria { get; set; }
    public string LessonEvaluation { get; set; }
    public string NextSteps { get; set; }
    public string Reflection { get; set; }
    public List<Step> Steps { get; set; }
}

public class Step
{
    public string LessonCode {get; set;} = "";
    public string Time { get; set; } = "";
    public string Stage { get; set; } = "";
    public string Differentiation { get; set; } = "";
    public string Management { get; set; } = "";
    public string StepType { get; set; } = "";
    public int Order { get; set; } = 0;
}


public class Lesson
{
    public string Code {get; set;}
    public string YearLevel { get; set; }
    public string Duration { get; set; }
    public string KlaTopic { get; set; }
    public string SchoolCode { get; set; }
    public string Curriculum { get; set; }
    public string PedagogicalStrategy { get; set; }
    public string ContentDescription { get; set; }
    public string LessonEvaluation { get; set; }
    public string NextSteps { get; set; }
    public string Reflection { get; set; }
    public bool Output { get; set; }
}


public class Learner
{
    public string LessonCode { get; set; }
    public string Name { get; set; }
    public string Differentiation { get; set; }
}


public class Elaboration
{
    public string LessonCode { get; set; }
    public string Name { get; set; }
    public string ElaborationCode { get; set; }
}

public class Outcome
{
    public string LessonCode { get; set; }
    public string Name { get; set; }
}

public class Resource
{
    public string LessonCode { get; set; }
    public string Name { get; set; }
}

public class Success
{
    public string LessonCode { get; set; }
    public string Name { get; set; }
}


public class LessonStageXl
{
    public string Code {get; set;}
    public string Time { get; set; }
    public string Stage { get; set; }
    public string Differentiation { get; set; }
    public string Management { get; set; }
}


