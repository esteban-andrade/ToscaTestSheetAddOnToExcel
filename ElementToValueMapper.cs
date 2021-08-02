using Tricentis.TCAPIObjects.Objects;

namespace TestSheetAddOn
{
    public static class ElementToValueMapper
    {
        public static string GetInstanceValueStringForElement(TDInstance topInstance, TDElementWrapper tdElementWrapper)
        {
            string instValueString = null;
            object instValue = GetInstanceValueForElement(topInstance, tdElementWrapper.ParentWrapper, tdElementWrapper);
            if (instValue != null)
            {
                instValueString = instValue.ToString();
                if (instValue is TDInstance)
                {
                    instValueString = (instValue as TDInstance).Name;
                }
            }

            return instValueString;
        }

        public static string GetInstanceValueStringForElementPreCondition(TDInstance topInstance, TDElementWrapper tdElementWrapper)
        {
            string instValueString = null;
            object instValue = GetInstanceValueForElementPrecondition(topInstance, tdElementWrapper.ParentWrapper, tdElementWrapper);
            if (instValue != null)
            {
                instValueString = instValue.ToString();
                if (instValue is TDInstance)
                {
                    instValueString = (instValue as TDInstance).Name;
                }
            }

            return instValueString;
        }

        private static object GetInstanceValueForElement(TDInstance topInstance, TDElementWrapper nextLevelUpWrapper, TDElementWrapper tdElementWrapper)
        {
            if (nextLevelUpWrapper == null)
            {
                object instValue = GetValueObject(topInstance, tdElementWrapper.TDElement);
                return instValue;
            }

            TDStructElement nextLevelUpElement = nextLevelUpWrapper.TDElement;
            TDClass nextLevelUpReferencedClass = (nextLevelUpElement as TDAttribute) == null ? null : (nextLevelUpElement as TDAttribute).ReferencedClass;
            TDInstances nextLevelUpInstances = (nextLevelUpReferencedClass != null) ? nextLevelUpReferencedClass.Instances : nextLevelUpElement.Instances;

            if (nextLevelUpInstances != null)
            {
                object nextLevelValue = GetInstanceValueForElement(topInstance, nextLevelUpWrapper.ParentWrapper, nextLevelUpWrapper);
                TDInstance nextLevelInstance = nextLevelValue as TDInstance;
                if (nextLevelInstance != null)
                {
                    return GetValueObject(nextLevelInstance, tdElementWrapper.TDElement);
                }
                return nextLevelValue;
            }
            return GetInstanceValueForElement(topInstance, nextLevelUpWrapper.ParentWrapper, tdElementWrapper);
        }

        private static object GetInstanceValueForElementPrecondition(TDInstance topInstance, TDElementWrapper nextLevelUpWrapper, TDElementWrapper tdElementWrapper)
        {
            if (nextLevelUpWrapper == null)
            {
                object instValue = GetValueObjectPrecondition(topInstance, tdElementWrapper.TDElement);
                return instValue;
            }

            TDStructElement nextLevelUpElement = nextLevelUpWrapper.TDElement;
            TDClass nextLevelUpReferencedClass = (nextLevelUpElement as TDAttribute) == null ? null : (nextLevelUpElement as TDAttribute).ReferencedClass;
            TDInstances nextLevelUpInstances = (nextLevelUpReferencedClass != null) ? nextLevelUpReferencedClass.Instances : nextLevelUpElement.Instances;

            if (nextLevelUpInstances != null)
            {
                object nextLevelValue = GetInstanceValueForElementPrecondition(topInstance, nextLevelUpWrapper.ParentWrapper, nextLevelUpWrapper);
                TDInstance nextLevelInstance = nextLevelValue as TDInstance;
                if (nextLevelInstance != null)
                {
                    return GetValueObjectPrecondition(nextLevelInstance, tdElementWrapper.TDElement);
                }
                return nextLevelValue;
            }
            return GetInstanceValueForElementPrecondition(topInstance, nextLevelUpWrapper.ParentWrapper, tdElementWrapper);
        }

        private static object GetValueObject(TDInstance instance, TDStructElement element)
        {
            foreach (TDInstanceValue v in instance.Values)
            {
                if (v.Element.UniqueId == element.UniqueId)
                {
                    if (v.ValueInstance != null)
                    {
                        return v.ValueInstance;
                    }
                    return v.Value;
                }
            }
            return null;
        }

        private static object GetValueObjectPrecondition(TDInstance instance, TDStructElement element)
        {
            foreach (TDInstanceValue v in instance.Values)
            {
                if (v.Element.UniqueId == element.UniqueId)
                {
                    if (v.ValueInstance != null)
                    {
                        return v.ValueInstance.Name;
                    }
                    return v.Value;
                }
            }
            return null;
        }

        private static object GetValueObjectPreconditionProcess(TDInstance instance, TDStructElement element)
        {
            foreach (TDInstanceValue v in instance.Values)
            {
                if (v.Element.UniqueId == element.UniqueId)
                {
                    if (v.ValueInstance != null)
                    {
                        return v.ValueInstance.Name;
                    }
                    return v.Value;
                }
            }
            return null;
        }


    }

}
