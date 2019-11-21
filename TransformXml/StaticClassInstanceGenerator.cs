using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using Microsoft.CSharp;
using System.CodeDom;
using System.CodeDom.Compiler;

namespace AndrewTweddle.Tools.TransformXml
{
    public static class StaticClassInstanceGenerator
    {
        public static object CreateCodeDOMObjectWrapperAroundStaticClass(
            Type staticClass, TextWriter dynamicCodeWriter,
            List<string> assemblyReferences)
        {
            string wrapperClassName = staticClass.Name + "Wrapper";
            string fullyQualifiedWrapperClassName
                = "AndrewTweddle.Tools.DynamicAssemblies." + wrapperClassName;

            CodeCompileUnit ccu = new CodeCompileUnit();
            CodeNamespace dynamicNamespace
                = new CodeNamespace("AndrewTweddle.Tools.DynamicAssemblies");
            ccu.Namespaces.Add(dynamicNamespace);

            CodeTypeDeclaration wrapperClass = new CodeTypeDeclaration(
                wrapperClassName);
            dynamicNamespace.Types.Add(wrapperClass);

            wrapperClass.Attributes = MemberAttributes.Public;
            wrapperClass.IsClass = true;

            CodeConstructor defaultCtor = new CodeConstructor();
            defaultCtor.Attributes = MemberAttributes.Public;
            wrapperClass.Members.Add(defaultCtor);

            /* Add wrapper methods for each public static method: */
            foreach (MethodInfo methInfo in staticClass.GetMethods())
            {
                if (methInfo.IsStatic && methInfo.IsPublic
                    && !methInfo.IsConstructor && !methInfo.IsGenericMethod
                    && !methInfo.IsAbstract)
                {
                    CreateCodeDOMWrapperAroundPublicStaticMethod(staticClass,
                        wrapperClass, methInfo);
                }
            }

            /* Compile the code: */
            CSharpCodeProvider provider = new CSharpCodeProvider();

            CodeGeneratorOptions codeGenOptions = new CodeGeneratorOptions();
            codeGenOptions.IndentString = "    ";
            codeGenOptions.ElseOnClosing = false;
            codeGenOptions.BracingStyle = "C";
            codeGenOptions.BlankLinesBetweenMembers = true;
            codeGenOptions.VerbatimOrder = true;

            dynamicCodeWriter.WriteLine("---------- ( {0} ) ----------",
                fullyQualifiedWrapperClassName);

            provider.GenerateCodeFromCompileUnit(ccu,
                dynamicCodeWriter, codeGenOptions);

            dynamicCodeWriter.WriteLine(
                "----------------------------------------------------");
            dynamicCodeWriter.WriteLine();

            CompilerParameters compilerOptions = new CompilerParameters();
            compilerOptions.GenerateInMemory = true;

            compilerOptions.ReferencedAssemblies.Add(
                staticClass.Module.Assembly.Location);

            /* Add references to other required assemblies: */
            foreach (string assemblyRef in assemblyReferences)
            {
                compilerOptions.ReferencedAssemblies.Add(assemblyRef);
            }

            CompilerResults compResults
                = provider.CompileAssemblyFromDom(compilerOptions, ccu);

            if (compResults.Errors.HasErrors)
            {
                StringBuilder compilationErrorsSB = new StringBuilder();
                StringWriter compilationErrorsSW 
                    = new StringWriter(compilationErrorsSB);
                compilationErrorsSW.WriteLine("Compilation errors for a "
                    + "dynamically generated assembly...");
                foreach (CompilerError err in compResults.Errors)
                {
                    compilationErrorsSW.WriteLine("{0}", err.ToString());
                }
                compilationErrorsSW.Flush();
                throw new Exception(compilationErrorsSB.ToString());
            }

            Type wrapperType
                = compResults.CompiledAssembly.GetType(
                    fullyQualifiedWrapperClassName);

            return Activator.CreateInstance(wrapperType);
        }

        private static void CreateCodeDOMWrapperAroundPublicStaticMethod(
            Type staticClass, CodeTypeDeclaration wrapperClass,
            MethodInfo methInfo)
        {
            ParameterInfo[] parameters = methInfo.GetParameters();
            int paramCount = parameters.Length;

            CodeMemberMethod wrappedMethod = new CodeMemberMethod();
            wrappedMethod.Name = methInfo.Name;
            wrappedMethod.Attributes
                = MemberAttributes.Public | MemberAttributes.Final;

            wrapperClass.Members.Add(wrappedMethod);

            CodeTypeReference returnTypeReference
                = new CodeTypeReference(methInfo.ReturnType);
            wrappedMethod.ReturnType = returnTypeReference;

            CodeMethodInvokeExpression methodInvokeExpression
                = new CodeMethodInvokeExpression();

            /* Determine whether the method has a return value
             * and decide whether to have a return statement
             * or just a method call statement in the method body:
             */
            CodeStatement methodBodyStatement;

            if (methInfo.ReturnType != typeof(void))
            {
                CodeMethodReturnStatement returnStatement
                     = new CodeMethodReturnStatement(
                         methodInvokeExpression);
                methodBodyStatement = returnStatement;
            }
            else
            {
                CodeExpressionStatement methodInvokeStatement
                    = new CodeExpressionStatement(
                        methodInvokeExpression);
                methodBodyStatement = methodInvokeStatement;
            }

            wrappedMethod.Statements.Add(methodBodyStatement);

            /* Now build up the parameters to the method: */
            for (int i = 0; i < paramCount; i++)
            {
                ParameterInfo paramInfo = parameters[i];

                CreateCodeDOMParameterForMethod(wrappedMethod,
                    methodInvokeExpression, paramInfo);
            }

            CodeTypeReferenceExpression staticClassReference
                = new CodeTypeReferenceExpression(staticClass);

            CodeMethodReferenceExpression methodReference
                = new CodeMethodReferenceExpression(
                    staticClassReference, methInfo.Name);

            methodInvokeExpression.Method = methodReference;
        }

        private static void CreateCodeDOMParameterForMethod(
            CodeMemberMethod wrappedMethod,
            CodeMethodInvokeExpression methodInvokeExpression,
            ParameterInfo paramInfo)
        {
            CodeParameterDeclarationExpression paramDeclaration;

            if (paramInfo.ParameterType.IsByRef)
            {
                paramDeclaration
                    = new CodeParameterDeclarationExpression(
                        paramInfo.ParameterType.GetElementType(),
                        paramInfo.Name);
                if (paramInfo.IsIn)
                {
                    paramDeclaration.Direction = FieldDirection.Ref;
                }
                else
                {
                    paramDeclaration.Direction = FieldDirection.Out;
                }
            }
            else
            {
                paramDeclaration
                    = new CodeParameterDeclarationExpression(
                        paramInfo.ParameterType, paramInfo.Name);
                paramDeclaration.Direction = FieldDirection.In;
            }

            wrappedMethod.Parameters.Add(paramDeclaration);

            CodeTypeReference paramTypeReference
                = new CodeTypeReference(paramInfo.ParameterType);

            CodeArgumentReferenceExpression paramReference
                = new CodeArgumentReferenceExpression(
                    paramInfo.Name);
            CodeDirectionExpression paramReferenceWithDirection
                = new CodeDirectionExpression(
                    paramDeclaration.Direction, paramReference);

            methodInvokeExpression.Parameters.Add(
                paramReferenceWithDirection);
        }

        /* My first attempt, using IL, which was too ambitious
         * as I just didn't know enough about how IL works...
        public static object CreateILObjectWrapperAroundStaticClass(
            Type staticClass, TextWriter dynamicCodeWriter,
            List<string> assemblyReferences)
        {
            AssemblyName assName 
                = new AssemblyName(staticClass.Name + "StaticClassWrapper");
            AssemblyBuilder assBuilder
                = AppDomain.CurrentDomain.DefineDynamicAssembly(assName,
                    AssemblyBuilderAccess.Run);
            ModuleBuilder modBuilder 
                = assBuilder.DefineDynamicModule("MainModule");
            TypeBuilder wrapperBuilder 
                = modBuilder.DefineType(staticClass.Name + "Wrapper",
                    TypeAttributes.Class | TypeAttributes.Public);

            wrapperBuilder.DefineDefaultConstructor(MethodAttributes.Public);

            foreach (MethodInfo methInfo in staticClass.GetMethods())
            {
                if (methInfo.IsConstructor)
                {
                    continue;
                }

                if (methInfo.IsStatic && methInfo.IsPublic)
                {
                    MethodBuilder methBuilder = wrapperBuilder.DefineMethod(
                        methInfo.Name, 
                        methInfo.Attributes & ~MethodAttributes.Static);

                    ParameterInfo[] paramInfos = methInfo.GetParameters();
                    Type[] paramTypes = new Type[paramInfos.Length];
                    
                    for (int i=0; i<paramInfos.Length; i++)
                    {
                        paramTypes[i] = paramInfos[i].GetType();
                    }
                    
                    methBuilder.SetReturnType(methInfo.ReturnType);
                    methBuilder.SetParameters(paramTypes);

                    foreach (ParameterInfo paramInfo in paramInfos)
                    {
                        ParameterBuilder paramBuilder
                            = methBuilder.DefineParameter(paramInfo.Position,
                                paramInfo.Attributes, paramInfo.Name);
                    }

                    ILGenerator ilg = methBuilder.GetILGenerator();

                    /* Here's where I run into difficulties
                     * and decided to opt for CodeDOM instead...
                }
            }

            return null;
        }
        */
    }
}
