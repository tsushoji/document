namespace BitOperationWinApp.Models
{
    public abstract class ValueObject<T> where T : ValueObject<T>
    {
        /// <summary>
        /// 等しいか
        /// </summary>
        /// <param name="obj">比較値</param>
        /// <returns>等しい場合、true そうでない場合、false</returns>
        public override bool Equals(object obj)
        {
            var vo = obj as T;
            if (vo == null)
            {
                return false;
            }

            return EqualsCore(vo);
        }

        /// <summary>
        /// ==
        /// </summary>
        /// <param name="vo1">比較値1</param>
        /// <param name="vo2">比較値2</param>
        /// <returns>等しい場合、true そうでない場合、false</returns>
        public static bool operator ==(ValueObject<T> vo1,
            ValueObject<T> vo2)
        {
            return Equals(vo1, vo2);
        }

        /// <summary>
        /// !=
        /// </summary>
        /// <param name="vo1">比較値1</param>
        /// <param name="vo2">比較値2</param>
        /// <returns>等しくない場合、true そうでない場合、false</returns>
        public static bool operator !=(ValueObject<T> vo1,
            ValueObject<T> vo2)
        {
            return !Equals(vo1, vo2);
        }
        protected abstract bool EqualsCore(T other);
        protected abstract int GetHashCodeCore();

        /// <summary>
        /// 文字列に変換
        /// </summary>
        /// <returns>変換した文字列</returns>
        public override string ToString()
        {
            return base.ToString();
        }

        /// <summary>
        /// ハッシュコード取得
        /// </summary>
        /// <returns>ハッシュコード</returns>
        public override int GetHashCode()
        {
            return GetHashCodeCore();
        }
    }
}
