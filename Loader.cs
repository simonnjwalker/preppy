#pragma warning disable CS8981

public class Loader
{

    public static List<Plan> GetLessonPlansFromXlsx(string filename)
    {
        List<Plan> plans = new List<Plan>();
        var xlsx = new Seamlex.Utilities.ExcelToData();
        var ds = xlsx.ToDataSet(filename);
        var lessons = xlsx.ToListData<Lesson>(ds.Tables["Lesson"]);
        var learners = xlsx.ToListData<Learner>(ds.Tables["Learner"]);
        var steps = xlsx.ToListData<Step>(ds.Tables["Step"]);
        var elaborations = xlsx.ToListData<Elaboration>(ds.Tables["Elaboration"]);
        var outcomes = xlsx.ToListData<Outcome>(ds.Tables["Outcome"]);
        var successes = xlsx.ToListData<Success>(ds.Tables["Success"]);
        var resources = xlsx.ToListData<Resource>(ds.Tables["Resource"]);

        // if at least one item is set in the 'output' field then only add those
        if(lessons.Count(x => !x.Output.Equals(true)) == 0)
            foreach (var lesson in lessons)
                lesson.Output = true;

        foreach(var lesson in lessons.Where(x => !x.Code.Equals("") && x.Output.Equals(true) ))
        {
            var plan = new Plan
                {
                    LessonCode = lesson.Code,
                    YearLevel = lesson.YearLevel,
                    Duration =lesson.Duration,
                    KlaTopic = lesson.KlaTopic,
                    SchoolCode = lesson.SchoolCode,
                    Curriculum = lesson.Curriculum,
                    PedagogicalStrategy = lesson.PedagogicalStrategy,
                    ContentDescription = lesson.ContentDescription,
                    LessonEvaluation = lesson.LessonEvaluation,
                    NextSteps = lesson.NextSteps,
                    Reflection = lesson.Reflection
                };
            plan.Learners = new List<Learner>();
            plan.Steps = new List<Step>();
            plan.Elaborations = new List<Elaboration>();
            plan.LearningOutcomes = new List<Outcome>();
            plan.SuccessCriteria = new List<Success>();
            plan.Resources = new List<Resource>();
            plan.Learners.AddRange(learners.Where(x => (x.LessonCode??"").Equals(plan.LessonCode)));
            plan.Steps.AddRange(steps.Where(x => (x.LessonCode??"").Equals(plan.LessonCode)));
            plan.Elaborations.AddRange(elaborations.Where(x => (x.LessonCode??"").Equals(plan.LessonCode)));
            plan.LearningOutcomes.AddRange(outcomes.Where(x => (x.LessonCode??"").Equals(plan.LessonCode)));
            plan.SuccessCriteria.AddRange(successes.Where(x => (x.LessonCode??"").Equals(plan.LessonCode)));
            plan.Resources.AddRange(resources.Where(x => (x.LessonCode??"").Equals(plan.LessonCode)));
            plans.Add(plan);
        }
        return plans;
    }
}
